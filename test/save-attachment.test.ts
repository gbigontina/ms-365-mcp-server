import { describe, expect, it, vi } from 'vitest';
import {
  CHUNK_SIZE,
  DEFAULT_FOLDER,
  SIMPLE_UPLOAD_LIMIT,
  chunkRanges,
  encodeDrivePath,
  safeFilename,
  saveAttachmentToOneDrive,
  simpleUploadPath,
  uploadSessionPath,
  validateFolder,
} from '../src/lib/save-attachment.js';

// The property under test: the attachment's bytes go from Graph to Graph inside the server, and the
// tool's answer carries metadata only. Every check below either holds that or holds the two upload
// paths Graph requires (single PUT up to 4 MB, upload session in 320 KiB multiples above it).

function fileAttachment(bytes: Buffer, over: Partial<Record<string, unknown>> = {}) {
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: 'INV289085.PDF',
    contentType: 'application/pdf',
    size: bytes.byteLength,
    contentBytes: bytes.toString('base64'),
    ...over,
  };
}

function fakeGraph(meta: unknown, driveItem: unknown = { id: 'item1', name: 'INV289085.PDF' }) {
  const calls: Array<{ endpoint: string; options: Record<string, unknown> }> = [];
  const graph = {
    async makeRequest(endpoint: string, options: Record<string, unknown> = {}) {
      calls.push({ endpoint, options });
      if (endpoint.includes('/attachments/')) return meta;
      if (endpoint.includes('createUploadSession'))
        return { uploadUrl: 'https://up.example/session' };
      return driveItem;
    },
  };
  return { graph, calls };
}

describe('paths and names', () => {
  it('encodes each folder segment and keeps the slashes', () => {
    expect(encodeDrivePath('/Documents/Claude/Projects/Personal Assistant/Attachments')).toBe(
      '/Documents/Claude/Projects/Personal%20Assistant/Attachments'
    );
    expect(encodeDrivePath('/a/b/')).toBe('/a/b');
  });

  it('builds the simple upload path with rename-on-conflict', () => {
    expect(simpleUploadPath('/x y', 'a b.pdf')).toBe(
      '/me/drive/root:/x%20y/a%20b.pdf:/content?@microsoft.graph.conflictBehavior=rename'
    );
    expect(uploadSessionPath('/x', 'a.pdf')).toBe('/me/drive/root:/x/a.pdf:/createUploadSession');
  });

  it('refuses a relative folder and dot segments', () => {
    expect(validateFolder('Documents')).toMatch(/starting with "\/"/);
    expect(validateFolder('/Documents/../Secrets')).toMatch(/\.\./);
    expect(validateFolder(DEFAULT_FOLDER)).toBeNull();
  });

  it('makes a filename OneDrive accepts', () => {
    expect(safeFilename('L 09-09-26.pdf')).toBe('L 09-09-26.pdf');
    expect(safeFilename('../a/b:c?.pdf')).toBe('_a_b_c_.pdf');
    expect(safeFilename('   ')).toBe('attachment');
  });

  it('splits an upload into 320 KiB multiples that cover every byte exactly once', () => {
    expect(CHUNK_SIZE % (320 * 1024)).toBe(0);
    const size = 2 * CHUNK_SIZE + 17;
    const ranges = chunkRanges(size);
    expect(ranges).toEqual([
      [0, CHUNK_SIZE - 1],
      [CHUNK_SIZE, 2 * CHUNK_SIZE - 1],
      [2 * CHUNK_SIZE, size - 1],
    ]);
    expect(chunkRanges(0)).toEqual([]);
  });
});

describe('saveAttachmentToOneDrive', () => {
  it('small file: one metadata GET, one PUT with the raw bytes, and a metadata-only answer', async () => {
    const bytes = Buffer.from('%PDF-1.4 hello');
    const { graph, calls } = fakeGraph(fileAttachment(bytes), {
      id: 'i1',
      name: 'INV289085.PDF',
      size: bytes.byteLength,
      webUrl: 'https://1drv/x',
      parentReference: {
        path: '/drive/root:/Documents/Claude/Projects/Personal Assistant/Attachments',
      },
    });
    const out = await saveAttachmentToOneDrive(graph, {
      messageId: 'M=1',
      attachmentId: 'A/2',
      folder: DEFAULT_FOLDER,
    });
    expect(calls).toHaveLength(2);
    expect(calls[0].endpoint).toBe('/me/messages/M%3D1/attachments/A%2F2');
    expect(calls[1].endpoint).toBe(simpleUploadPath(DEFAULT_FOLDER, 'INV289085.PDF'));
    expect(calls[1].options.method).toBe('PUT');
    expect(Buffer.isBuffer(calls[1].options.body)).toBe(true);
    expect((calls[1].options.body as Buffer).equals(bytes)).toBe(true);
    expect((calls[1].options.headers as Record<string, string>)['Content-Type']).toBe(
      'application/pdf'
    );
    expect(out.method).toBe('put');
    expect(out.bytesUploaded).toBe(bytes.byteLength);
    expect(out.driveItem.id).toBe('i1');
    expect(JSON.stringify(out)).not.toContain(bytes.toString('base64'));
  });

  it('large file: an upload session, every chunk PUT to the session url with a Content-Range', async () => {
    const size = SIMPLE_UPLOAD_LIMIT + CHUNK_SIZE + 5;
    const bytes = Buffer.alloc(size, 7);
    const { graph, calls } = fakeGraph(fileAttachment(bytes, { name: 'Capitolato poggioli.pdf' }));
    const puts: Array<{ url: string; range: string; len: number }> = [];
    const fetchImpl = vi.fn(async (url: string, init: { headers?: unknown; body?: unknown }) => {
      const headers = init.headers as Record<string, string>;
      puts.push({ url, range: headers['Content-Range'], len: (init.body as Buffer).byteLength });
      const last = puts.length === chunkRanges(size).length;
      return {
        ok: true,
        status: last ? 201 : 202,
        text: async () =>
          last ? JSON.stringify({ id: 'big1', name: 'Capitolato poggioli.pdf', size }) : '',
      } as unknown as Response;
    });
    const out = await saveAttachmentToOneDrive(
      graph,
      { messageId: 'm', attachmentId: 'a', folder: '/Docs' },
      fetchImpl as unknown as typeof fetch
    );
    expect(calls.map((c) => c.endpoint)).toEqual([
      '/me/messages/m/attachments/a',
      uploadSessionPath('/Docs', 'Capitolato poggioli.pdf'),
    ]);
    expect(puts.every((p) => p.url === 'https://up.example/session')).toBe(true);
    expect(puts.map((p) => p.range)).toEqual(
      chunkRanges(size).map(([s, e]) => `bytes ${s}-${e}/${size}`)
    );
    expect(puts.reduce((n, p) => n + p.len, 0)).toBe(size);
    expect(out.method).toBe('session');
    expect(out.driveItem.id).toBe('big1');
  });

  it('refuses an itemAttachment instead of saving something that is not a file', async () => {
    const { graph, calls } = fakeGraph({
      '@odata.type': '#microsoft.graph.itemAttachment',
      name: 'fwd',
    });
    await expect(
      saveAttachmentToOneDrive(graph, { messageId: 'm', attachmentId: 'a', folder: '/Docs' })
    ).rejects.toThrow(/itemAttachment/);
    expect(calls).toHaveLength(1);
  });

  it('refuses a bad folder before touching Graph', async () => {
    const { graph, calls } = fakeGraph(fileAttachment(Buffer.from('x')));
    await expect(
      saveAttachmentToOneDrive(graph, { messageId: 'm', attachmentId: 'a', folder: 'Docs' })
    ).rejects.toThrow(/starting with/);
    expect(calls).toHaveLength(0);
  });

  it('passes the per-account token through to every Graph call', async () => {
    const { graph, calls } = fakeGraph(fileAttachment(Buffer.from('x')));
    await saveAttachmentToOneDrive(graph, {
      messageId: 'm',
      attachmentId: 'a',
      folder: '/Docs',
      accessToken: 'tok',
    });
    expect(calls.every((c) => c.options.accessToken === 'tok')).toBe(true);
  });
});
