import { NextRequest, NextResponse } from 'next/server';

export const maxDuration = 30;

const BASE = 'https://live.perigeeportal.co.za';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

// Extract value of a named field from HTML
function extractField(html: string, name: string): string {
  const match = html.match(new RegExp(`name="${name}"[^>]*value="([^"]*)"`, 'i'))
    || html.match(new RegExp(`value="([^"]*)"[^>]*name="${name}"`, 'i'));
  return match?.[1] ?? '';
}

// Parse Set-Cookie headers into a single Cookie header string
function cookieString(setCookieHeaders: string[]): string {
  return setCookieHeaders.map(c => c.split(';')[0]).filter(Boolean).join('; ');
}

export async function POST(req: NextRequest) {
  try {
    const { username, password } = await req.json();
    if (!username || !password) {
      return NextResponse.json({ error: 'Username and password required' }, { status: 400 });
    }

    // ── Step 1: GET login page to obtain CSRF tokens ──────────────────────────
    const loginPageRes = await fetch(`${BASE}/user/login`, {
      headers: { 'User-Agent': UA, 'Accept': 'text/html' },
    });
    if (!loginPageRes.ok) {
      return NextResponse.json({ error: `Could not reach Perigee portal (${loginPageRes.status})` }, { status: 502 });
    }

    const html = await loginPageRes.text();
    const formBuildId = extractField(html, 'form_build_id');
    const formId = extractField(html, 'form_id') || 'user_login_form';

    const initCookies = loginPageRes.headers.getSetCookie?.() ?? [];

    // ── Step 2: POST credentials ──────────────────────────────────────────────
    const body = new URLSearchParams({
      name: username,
      pass: password,
      form_build_id: formBuildId,
      form_id: formId,
      op: 'Log in',
    });

    const loginRes = await fetch(`${BASE}/user/login`, {
      method: 'POST',
      redirect: 'manual',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': UA,
        'Referer': `${BASE}/user/login`,
        'Cookie': cookieString(initCookies),
      },
      body: body.toString(),
    });

    // Drupal redirects (302) on success, stays on login page (200) on failure
    const isRedirect = loginRes.status === 302 || loginRes.status === 303;
    const sessionCookies = loginRes.headers.getSetCookie?.() ?? [];
    const allCookies = [...initCookies, ...sessionCookies];
    const sessionStr = cookieString(allCookies);

    if (!isRedirect || !sessionStr) {
      return NextResponse.json({ error: 'Invalid Perigee credentials. Please try again.' }, { status: 401 });
    }

    // ── Step 3: Follow redirect to confirm login and get username ─────────────
    const redirectUrl = loginRes.headers.get('location') || `${BASE}/user`;
    const profileRes = await fetch(
      redirectUrl.startsWith('http') ? redirectUrl : `${BASE}${redirectUrl}`,
      { headers: { 'User-Agent': UA, 'Cookie': sessionStr } }
    );
    const profileHtml = await profileRes.text();

    // Try to extract display name from Drupal profile page
    const nameMatch = profileHtml.match(/<h1[^>]*class="[^"]*page-header[^"]*"[^>]*>([^<]+)<\/h1>/i)
      || profileHtml.match(/<title>([^|<]+)/i);
    const displayName = nameMatch?.[1]?.trim().replace(/ \|.*$/, '') ?? username;

    // ── Step 4: Store session cookie in our httpOnly cookie ───────────────────
    const response = NextResponse.json({ success: true, displayName });
    response.cookies.set('perigee_session', sessionStr, {
      httpOnly: true,
      secure: true,
      sameSite: 'strict',
      maxAge: 60 * 60 * 8, // 8 hours
      path: '/',
    });
    return response;

  } catch (err) {
    console.error('Perigee login error:', err);
    return NextResponse.json({ error: 'Login failed. Please try again.' }, { status: 500 });
  }
}

// Logout — clears the session cookie
export async function DELETE() {
  const response = NextResponse.json({ success: true });
  response.cookies.delete('perigee_session');
  return response;
}
