import { getStore } from '@netlify/blobs';

export default async (req: Request) => {
  const url = new URL(req.url);
  const key = url.searchParams.get('key');
  const name = url.searchParams.get('name') || 'report.docx';

  if (!key) return new Response('Missing key', { status: 400 });

  const store = getStore({ name: 'docx-temp', consistency: 'strong' });
  const blob = await store.get(key, { type: 'arrayBuffer' });

  if (!blob) return new Response('File not found or expired', { status: 404 });

  return new Response(blob, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(name)}`,
    },
  });
};

export const config = { path: '/api/download-docx' };
