import { NextRequest } from 'next/server';

export const runtime = 'edge';
// Run on non-US regions — Perigee's firewall likely only blocks US cloud IPs
export const preferredRegion = ['gru1', 'sin1', 'syd1']; // Sao Paulo, Singapore, Sydney

const ALLOWED_ORIGIN = 'https://live.perigeeportal.co.za';

export async function GET(req: NextRequest) {
  const url = req.nextUrl.searchParams.get('url');

  if (!url || !url.startsWith(ALLOWED_ORIGIN + '/')) {
    return new Response('Bad request', { status: 400 });
  }

  try {
    const res = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'image/jpeg,image/png,image/webp,image/*,*/*;q=0.8',
        'Referer': ALLOWED_ORIGIN + '/',
      },
    });

    if (!res.ok) {
      return new Response(`Upstream returned ${res.status}`, { status: res.status });
    }

    const contentType = res.headers.get('Content-Type') || 'image/jpeg';

    return new Response(res.body, {
      status: 200,
      headers: {
        'Content-Type': contentType,
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'public, max-age=3600',
      },
    });
  } catch (err) {
    console.error('Proxy error:', err);
    return new Response('Proxy fetch failed', { status: 502 });
  }
}
