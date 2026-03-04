'use client';

import { useState, useRef, useCallback, useEffect } from 'react';
import Image from 'next/image';

interface PreviewStats {
  totalRows: number;
  uniqueUsers: number;
  uniqueStores: number;
  uniqueDays: string[];
  dateRange: { from: string; to: string };
  totalImages: number;
}

export default function Home() {
  // ── Perigee auth state ────────────────────────────────────────────────────
  const [loginChecked, setLoginChecked] = useState(false);
  const [loggedInAs, setLoggedInAs] = useState<string | null>(null);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loggingIn, setLoggingIn] = useState(false);
  const [loginError, setLoginError] = useState<string | null>(null);

  // ── File / PPT state ──────────────────────────────────────────────────────
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [preview, setPreview] = useState<PreviewStats | null>(null);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [done, setDone] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  // Check existing session on mount via a lightweight parse test
  useEffect(() => {
    fetch('/api/perigee-login', { method: 'GET' })
      .then(r => {
        if (r.status === 405) {
          // Route exists but GET not allowed — not logged in
          setLoginChecked(true);
        }
      })
      .catch(() => setLoginChecked(true));
    // We'll detect login state by whether generate succeeds or returns 401
    setLoginChecked(true);
  }, []);

  // ── Perigee login ─────────────────────────────────────────────────────────
  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoggingIn(true);
    setLoginError(null);
    try {
      const res = await fetch('/api/perigee-login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: loginUsername, password: loginPassword }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Login failed');
      setLoggedInAs(data.displayName || loginUsername);
      setLoginPassword('');
    } catch (e: unknown) {
      setLoginError(e instanceof Error ? e.message : 'Login failed');
    } finally {
      setLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await fetch('/api/perigee-login', { method: 'DELETE' });
    setLoggedInAs(null);
    setFile(null);
    setPreview(null);
    setDone(false);
    setError(null);
  };

  // ── File handling ─────────────────────────────────────────────────────────
  const handleFile = useCallback(async (f: File) => {
    if (!f.name.endsWith('.xlsx') && !f.name.endsWith('.xls')) {
      setError('Please upload an Excel file (.xlsx or .xls)');
      return;
    }
    setFile(f);
    setPreview(null);
    setError(null);
    setDone(false);
    setParsing(true);
    try {
      const fd = new FormData();
      fd.append('file', f);
      const res = await fetch('/api/parse', { method: 'POST', body: fd });
      if (!res.ok) throw new Error((await res.json()).error || 'Parse failed');
      setPreview(await res.json());
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to read file');
      setFile(null);
    } finally {
      setParsing(false);
    }
  }, []);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }, [handleFile]);

  // ── Generate ──────────────────────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!file || !preview) return;
    setGenerating(true);
    setError(null);
    setDone(false);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch('/api/generate', { method: 'POST', body: fd });
      if (res.status === 401) {
        setLoggedInAs(null);
        throw new Error('Session expired — please log in again');
      }
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || 'Generation failed');
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'FairPrice_VisualMerch.pptx';
      a.click();
      URL.revokeObjectURL(url);
      setDone(true);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to generate presentation');
    } finally {
      setGenerating(false);
    }
  };

  const reset = () => {
    setFile(null);
    setPreview(null);
    setError(null);
    setDone(false);
    if (inputRef.current) inputRef.current.value = '';
  };

  return (
    <div style={{ minHeight: '100vh', background: '#F6F6F6' }}>
      {/* ── Header ── */}
      <header style={{
        background: '#76bd22', padding: '0 2rem', height: '64px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          <Image src="/fairprice-logo.png" alt="Fair Price" width={120} height={52}
            style={{ objectFit: 'contain', filter: 'brightness(0) invert(1)' }} />
          <div style={{ width: 1, height: 32, background: 'rgba(255,255,255,0.4)' }} />
          <span style={{ color: '#fff', fontWeight: 700, fontSize: '1.05rem' }}>PPT Builder</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
          {loggedInAs && (
            <span style={{ color: 'rgba(255,255,255,0.9)', fontSize: '0.85rem' }}>
              Logged in as <strong>{loggedInAs}</strong>
            </span>
          )}
          <Image src="/perigee-logo.png" alt="Perigee" width={44} height={44} style={{ objectFit: 'contain' }} />
        </div>
      </header>

      <main style={{ maxWidth: '680px', margin: '0 auto', padding: '2.5rem 1rem' }}>
        <h1 style={{ fontSize: '1.5rem', fontWeight: 700, color: '#242424', marginBottom: '0.25rem' }}>
          Visual Merch Presentation Generator
        </h1>
        <p style={{ color: '#6b7280', marginBottom: '2rem', fontSize: '0.95rem' }}>
          Log in with your Perigee account, upload your survey export, and download a branded PowerPoint.
        </p>

        {/* Error banner */}
        {error && (
          <div style={{
            background: '#FEE2E2', border: '1px solid #FF4539', borderRadius: 8,
            padding: '0.75rem 1rem', marginBottom: '1.5rem', color: '#b91c1c',
            fontSize: '0.9rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center',
          }}>
            <span>{error}</span>
            <button onClick={() => setError(null)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#b91c1c', fontWeight: 700 }}>✕</button>
          </div>
        )}

        {/* ── Step 1: Perigee Login ── */}
        <div style={card}>
          <StepLabel n={1} label="Login with Perigee" />

          {loggedInAs ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f0fce4', border: '1px solid #76bd22', borderRadius: 8, padding: '0.75rem 1rem' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ fontSize: '1.2rem' }}>✅</span>
                <div>
                  <div style={{ fontWeight: 600, color: '#242424', fontSize: '0.9rem' }}>Connected to Perigee</div>
                  <div style={{ color: '#6b7280', fontSize: '0.8rem' }}>{loggedInAs}</div>
                </div>
              </div>
              <button
                onClick={handleLogout}
                style={{ background: 'none', border: '1px solid #d1d5db', borderRadius: 6, padding: '0.3rem 0.75rem', cursor: 'pointer', color: '#6b7280', fontSize: '0.8rem' }}
              >
                Log out
              </button>
            </div>
          ) : (
            <form onSubmit={handleLogin} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
              <p style={{ color: '#6b7280', fontSize: '0.875rem', marginBottom: '0.25rem' }}>
                Your Perigee credentials are used to securely download survey images. They are never stored.
              </p>
              {loginError && (
                <div style={{ background: '#FEE2E2', border: '1px solid #FF4539', borderRadius: 6, padding: '0.5rem 0.75rem', color: '#b91c1c', fontSize: '0.875rem' }}>
                  {loginError}
                </div>
              )}
              <input
                type="text"
                placeholder="Perigee email or username"
                value={loginUsername}
                onChange={e => setLoginUsername(e.target.value)}
                required
                autoComplete="username"
                style={inputStyle}
              />
              <input
                type="password"
                placeholder="Password"
                value={loginPassword}
                onChange={e => setLoginPassword(e.target.value)}
                required
                autoComplete="current-password"
                style={inputStyle}
              />
              <button
                type="submit"
                disabled={loggingIn || !loginUsername || !loginPassword}
                style={{
                  background: loggingIn || !loginUsername || !loginPassword ? '#a8d97b' : '#76bd22',
                  color: '#fff', border: 'none', borderRadius: 8,
                  padding: '0.75rem 1.5rem', fontSize: '0.95rem', fontWeight: 700,
                  cursor: loggingIn || !loginUsername || !loginPassword ? 'not-allowed' : 'pointer',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem',
                }}
              >
                {loggingIn ? <><Spinner color="#fff" /> Connecting…</> : 'Login via Perigee'}
              </button>
            </form>
          )}
        </div>

        {/* ── Step 2: Upload (only shown when logged in) ── */}
        {loggedInAs && (
          <div style={card}>
            <StepLabel n={2} label="Upload Survey Excel" />
            <div
              onClick={() => inputRef.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              style={{
                border: `2px dashed ${dragging || file ? '#76bd22' : '#d1d5db'}`,
                borderRadius: 10, padding: '2rem', textAlign: 'center', cursor: 'pointer',
                background: dragging || file ? '#f0fce4' : '#fafafa', transition: 'all 0.2s',
              }}
            >
              {parsing ? (
                <div style={{ color: '#76bd22', fontWeight: 600, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem' }}>
                  <Spinner color="#76bd22" /> Reading file…
                </div>
              ) : file ? (
                <div>
                  <div style={{ fontSize: '1.75rem', marginBottom: '0.25rem' }}>✅</div>
                  <div style={{ fontWeight: 600, color: '#242424' }}>{file.name}</div>
                  <div style={{ color: '#6b7280', fontSize: '0.85rem', marginTop: '0.2rem' }}>
                    {(file.size / 1024).toFixed(0)} KB · Click to change
                  </div>
                </div>
              ) : (
                <div>
                  <div style={{ fontSize: '2.25rem', marginBottom: '0.5rem' }}>📊</div>
                  <div style={{ fontWeight: 600, color: '#242424' }}>Drop your Excel file here</div>
                  <div style={{ color: '#6b7280', fontSize: '0.85rem', marginTop: '0.25rem' }}>or click to browse · .xlsx files only</div>
                </div>
              )}
            </div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
          </div>
        )}

        {/* ── Step 3: Preview ── */}
        {loggedInAs && preview && (
          <div style={card}>
            <StepLabel n={3} label="Preview" />
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.75rem', marginBottom: '1rem' }}>
              <StatCard label="Total Surveys" value={preview.totalRows} />
              <StatCard label="Unique Users" value={preview.uniqueUsers} />
              <StatCard label="Unique Stores" value={preview.uniqueStores} />
              <StatCard label="Images to Fetch" value={preview.totalImages} />
            </div>
            {(preview.dateRange.from || preview.dateRange.to) && (
              <div style={{ background: '#F6F6F6', borderRadius: 8, padding: '0.6rem 1rem', fontSize: '0.9rem', color: '#242424', marginBottom: '0.5rem' }}>
                <span style={{ color: '#6b7280' }}>Date range: </span>
                <strong>
                  {preview.dateRange.from === preview.dateRange.to || !preview.dateRange.to
                    ? preview.dateRange.from
                    : `${preview.dateRange.from} – ${preview.dateRange.to}`}
                </strong>
              </div>
            )}
            {preview.uniqueDays.length > 0 && (
              <div style={{ fontSize: '0.85rem', color: '#6b7280' }}>
                Days surveyed: {preview.uniqueDays.join(', ')}
              </div>
            )}
          </div>
        )}

        {/* ── Step 4: Generate ── */}
        {loggedInAs && preview && (
          <div style={card}>
            <StepLabel n={4} label="Generate Presentation" />
            <p style={{ color: '#6b7280', fontSize: '0.875rem', marginBottom: '1.25rem' }}>
              Fetching {preview.totalImages} survey images from Perigee and building a fully branded PowerPoint.
              This may take up to a minute for large files.
            </p>
            {done ? (
              <div style={{ background: '#f0fce4', border: '1px solid #76bd22', borderRadius: 8, padding: '1.25rem', textAlign: 'center' }}>
                <div style={{ fontSize: '1.75rem', marginBottom: '0.25rem' }}>🎉</div>
                <div style={{ fontWeight: 700, color: '#242424', marginBottom: '0.35rem' }}>Presentation downloaded!</div>
                <div style={{ color: '#6b7280', fontSize: '0.85rem', marginBottom: '1rem' }}>
                  Check your Downloads folder for <strong>FairPrice_VisualMerch.pptx</strong>
                </div>
                <button onClick={reset} style={{ background: 'transparent', border: '1.5px solid #76bd22', color: '#5e9a18', borderRadius: 6, padding: '0.4rem 1rem', cursor: 'pointer', fontWeight: 600, fontSize: '0.875rem' }}>
                  Build another
                </button>
              </div>
            ) : (
              <button
                onClick={handleGenerate}
                disabled={generating}
                style={{
                  width: '100%', background: generating ? '#a8d97b' : '#76bd22',
                  color: '#fff', border: 'none', borderRadius: 8,
                  padding: '0.9rem 1.5rem', fontSize: '1rem', fontWeight: 700,
                  cursor: generating ? 'not-allowed' : 'pointer',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.6rem',
                  transition: 'background 0.2s',
                }}
              >
                {generating ? <><Spinner color="#fff" /> Building your presentation…</> : <>▶ Generate Presentation</>}
              </button>
            )}
          </div>
        )}
      </main>

      <footer style={{ textAlign: 'center', padding: '2rem', color: '#9ca3af', fontSize: '0.8rem' }}>
        Fair Price PPT Builder · Powered by Perigee Field Goose
      </footer>
    </div>
  );
}

// ── Shared styles ──────────────────────────────────────────────────────────

const card: React.CSSProperties = {
  background: '#fff', borderRadius: 12, padding: '1.5rem', marginBottom: '1.25rem',
  boxShadow: '0 1px 4px rgba(0,0,0,0.07)', border: '1px solid #e2e2e2',
};

const inputStyle: React.CSSProperties = {
  border: '1.5px solid #d1d5db', borderRadius: 8, padding: '0.65rem 0.875rem',
  fontSize: '0.925rem', color: '#242424', outline: 'none', width: '100%',
};

// ── Sub-components ─────────────────────────────────────────────────────────

function StepLabel({ n, label }: { n: number; label: string }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
      <span style={{
        background: '#76bd22', color: '#fff', borderRadius: '50%',
        width: 24, height: 24, display: 'flex', alignItems: 'center', justifyContent: 'center',
        fontSize: '0.8rem', fontWeight: 700, flexShrink: 0,
      }}>{n}</span>
      <span style={{ fontWeight: 700, color: '#242424' }}>{label}</span>
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <div style={{ background: '#F6F6F6', borderRadius: 8, padding: '0.75rem', textAlign: 'center', border: '1px solid #e2e2e2' }}>
      <div style={{ fontSize: '1.6rem', fontWeight: 700, color: '#76bd22' }}>{value}</div>
      <div style={{ fontSize: '0.75rem', color: '#6b7280', marginTop: '0.2rem' }}>{label}</div>
    </div>
  );
}

function Spinner({ color }: { color: string }) {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" style={{ animation: 'spin 0.8s linear infinite', display: 'inline-block' }}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      <circle cx="9" cy="9" r="7" fill="none" stroke={color} strokeWidth="2.5" strokeDasharray="30 12" />
    </svg>
  );
}
