/**
 * save-attachment-to-onedrive — copy a mail attachment into the signed-in user's OneDrive
 * entirely server-side, so the bytes never pass through the MCP client (the model).
 *
 * Why this exists: `download-bytes` returns base64 into the model's context. That is fine for a
 * profile photo and useless for a 500 KB invoice — the content would have to travel through the
 * model's own output to become a file. This tool fetches the attachment with Graph and uploads it
 * with Graph, in the server process, and returns only metadata.
 *
 * Small files (≤ SIMPLE_UPLOAD_LIMIT) use a single PUT to `/content`; larger ones use an upload
 * session in CHUNK_SIZE pieces, which Graph requires to be multiples of 320 KiB.
 */

export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // Graph: single PUT allowed up to 4 MB
export const CHUNK_SIZE = 10 * 320 * 1024; // 3,276,800 bytes, a multiple of 320 KiB
export const DEFAULT_FOLDER = '/Documents/Claude/Projects/Personal Assistant/Attachments';

/** Encode a OneDrive folder path segment by segment, keeping the slashes. */
export function encodeDrivePath(folder: string): string {
  const trimmed = folder.replace(/\/+$/, '');
  return trimmed
    .split('/')
    .map((seg) => (seg === '' ? '' : encodeURIComponent(seg)))
    .join('/');
}

/** A filename OneDrive will accept: no path separators, no control characters, no leading dots. */
export function safeFilename(name: string): string {
  const cleaned = name
    .replace(/[\\/:*?"<>|]/g, '_')
    // eslint-disable-next-line no-control-regex
    .replace(/[\x00-\x1f]/g, '')
    .trim()
    .replace(/^\.+/, '');
  return cleaned.length > 0 ? cleaned : 'attachment';
}

/** Validate the folder parameter: absolute (leading slash), no `..`, no empty middle segments. */
export function validateFolder(folder: string): string | null {
  if (typeof folder !== 'string' || !folder.startsWith('/')) {
    return 'folder must be a OneDrive path starting with "/", e.g. ' + DEFAULT_FOLDER;
  }
  const segs = folder
    .split('/')
    .slice(1)
    .filter((s) => s !== '');
  if (segs.some((s) => s === '..' || s === '.')) {
    return 'folder must not contain "." or ".." segments';
  }
  return null;
}

/** `[start, end]` byte ranges (inclusive) for an upload session over `size` bytes. */
export function chunkRanges(size: number, chunk = CHUNK_SIZE): Array<[number, number]> {
  if (size <= 0) return [];
  const out: Array<[number, number]> = [];
  for (let start = 0; start < size; start += chunk) {
    out.push([start, Math.min(start + chunk, size) - 1]);
  }
  return out;
}

/** The Graph path for a simple upload, with rename-on-conflict. */
export function simpleUploadPath(folder: string, name: string): string {
  return `/me/drive/root:${encodeDrivePath(folder)}/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=rename`;
}

/** The Graph path that opens an upload session for a large file. */
export function uploadSessionPath(folder: string, name: string): string {
  return `/me/drive/root:${encodeDrivePath(folder)}/${encodeURIComponent(name)}:/createUploadSession`;
}

export interface AttachmentMeta {
  '@odata.type'?: string;
  name?: string;
  contentType?: string;
  size?: number;
  contentBytes?: string;
  isInline?: boolean;
}

export interface DriveItemSummary {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  parentReference?: { path?: string };
}

/** What the tool returns: metadata only, never the bytes. */
export function summarize(
  meta: AttachmentMeta,
  item: DriveItemSummary,
  folder: string,
  bytes: number,
  method: 'put' | 'session'
) {
  return {
    saved: true,
    attachment: { name: meta.name, contentType: meta.contentType, size: meta.size ?? bytes },
    driveItem: {
      id: item.id,
      name: item.name,
      size: item.size,
      webUrl: item.webUrl,
      parentPath: item.parentReference?.path,
    },
    folder,
    bytesUploaded: bytes,
    method,
  };
}

/** Minimal surface of GraphClient this module needs, so it can be tested with a fake. */
export interface GraphLike {
  makeRequest(endpoint: string, options?: Record<string, unknown>): Promise<unknown>;
}

export interface SaveAttachmentParams {
  messageId: string;
  attachmentId: string;
  folder: string;
  filename?: string;
  accessToken?: string;
}

/**
 * The whole operation. `fetchImpl` is injectable for tests (upload-session PUTs go to an absolute
 * URL Graph hands back, which the GraphClient cannot build, so they use fetch directly).
 */
export async function saveAttachmentToOneDrive(
  graph: GraphLike,
  params: SaveAttachmentParams,
  fetchImpl: typeof fetch = fetch
) {
  const folderError = validateFolder(params.folder);
  if (folderError) throw new Error(folderError);

  const meta = (await graph.makeRequest(
    `/me/messages/${encodeURIComponent(params.messageId)}/attachments/${encodeURIComponent(params.attachmentId)}`,
    { accessToken: params.accessToken }
  )) as AttachmentMeta;

  const kind = meta['@odata.type'] ?? '';
  if (kind !== '#microsoft.graph.fileAttachment') {
    throw new Error(
      `Only file attachments can be saved; this one is ${kind || 'of unknown type'} ` +
        '(an itemAttachment is an embedded message or event, a referenceAttachment is a link).'
    );
  }
  if (typeof meta.contentBytes !== 'string' || meta.contentBytes.length === 0) {
    throw new Error('Graph returned the attachment without contentBytes; nothing to save.');
  }

  const bytes = Buffer.from(meta.contentBytes, 'base64');
  const name = safeFilename(params.filename ?? meta.name ?? 'attachment');

  if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
    const item = (await graph.makeRequest(simpleUploadPath(params.folder, name), {
      method: 'PUT',
      body: bytes,
      headers: { 'Content-Type': meta.contentType || 'application/octet-stream' },
      accessToken: params.accessToken,
    })) as DriveItemSummary;
    return summarize(meta, item, params.folder, bytes.byteLength, 'put');
  }

  const session = (await graph.makeRequest(uploadSessionPath(params.folder, name), {
    method: 'POST',
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } }),
    accessToken: params.accessToken,
  })) as { uploadUrl?: string };
  if (!session.uploadUrl) throw new Error('Graph did not return an uploadUrl for the session.');

  let last: unknown = null;
  for (const [start, end] of chunkRanges(bytes.byteLength)) {
    const piece = bytes.subarray(start, end + 1);
    const res = await fetchImpl(session.uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(piece.byteLength),
        'Content-Range': `bytes ${start}-${end}/${bytes.byteLength}`,
      },
      // Node fetch accepts a Buffer; the lib DOM typing does not name it, and eslint has no BodyInit global.
      body: piece as unknown as string,
    });
    if (!res.ok) {
      throw new Error(
        `Upload session chunk ${start}-${end} failed: ${res.status} ${await res.text()}`
      );
    }
    const text = await res.text();
    last = text ? JSON.parse(text) : null;
  }
  return summarize(
    meta,
    (last ?? {}) as DriveItemSummary,
    params.folder,
    bytes.byteLength,
    'session'
  );
}
