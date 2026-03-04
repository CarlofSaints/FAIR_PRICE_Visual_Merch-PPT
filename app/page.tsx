'use client';

import { useState, useRef, useCallback, useEffect } from 'react';
import Image from 'next/image';
import type { ParsedDataJSON, SurveyRowJSON } from '@/lib/ppt-builder-browser';
import type { UserSummary } from '@/types/survey';

const PROXY_KEY = 'fp_cf_proxy_url';

interface PreviewStats {
  totalRows: number;
  uniqueUsers: string[];
  uniqueStores: string[];
  uniqueDays: string[];
  dateRange: { from: string; to: string };
  totalImages: number;
  rows: SurveyRowJSON[];
  userSummaries: UserSummary[];
}

export default function Home() {
  // ── Proxy config ────────────────────────────────────────────────────────────
  const [proxyUrl, setProxyUrl] = useState('');
  const [proxyInput, setProxyInput] = useState('');
  const [proxyOpen, setProxyOpen] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem(PROXY_KEY) || '';
    setProxyUrl(saved);
    setProxyInput(saved);
  }, []);

  const saveProxy = () => {
    const trimmed = proxyInput.trim();
    localStorage.setItem(PROXY_KEY, trimmed);
    setProxyUrl(trimmed);
    setProxyOpen(false);
  };

  // ── File / PPT state ─────────────────────────────────────────────────────
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [preview, setPreview] = useState<PreviewStats | null>(null);
  const [generating, setGenerating] = useState(false);
  const [imgProgress, setImgProgress] = useState<{ loaded: number; total: number } | null>(null);
  const [buildingPpt, setBuildingPpt] = useState(false);
  const [done, setDone] = useState<{ imagesLoaded: number; imagesTotal: number; httpErrors: number; exceptions: number; firstError: string; proxyUsed: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // ── File handling ───────────────────────────────────────────────────────────
  const handleFile = useCallback(async (f: File) => {
    if (!f.name.endsWith('.xlsx') && !f.name.endsWith('.xls')) {
      setError('Please upload an Excel file (.xlsx or .xls)');
      return;
    }
    setFile(f);
    setPreview(null);
    setError(null);
    setDone(null);
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

  // ── Generate (client-side) ──────────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!preview) return;
    setGenerating(true);
    setError(null);
    setDone(null);
    setImgProgress(null);
    setBuildingPpt(false);
    try {
      const { buildPptxBrowser, setProxyBase } = await import('@/lib/ppt-builder-browser');
      if (proxyUrl) setProxyBase(proxyUrl);

      const data: ParsedDataJSON = {
        rows: preview.rows,
        uniqueUsers: preview.uniqueUsers,
        uniqueStores: preview.uniqueStores,
        uniqueDays: preview.uniqueDays,
        dateRange: preview.dateRange,
        totalRows: preview.totalRows,
      };

      const result = await buildPptxBrowser(
        data,
        preview.userSummaries,
        (loaded, total) => {
          if (loaded === total && total > 0) {
            setBuildingPpt(true);
          }
          setImgProgress({ loaded, total });
        },
      );

      setBuildingPpt(false);

      const url = URL.createObjectURL(result.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'FairPrice_VisualMerch.pptx';
      a.click();
      URL.revokeObjectURL(url);
      setDone({
        imagesLoaded: result.imagesLoaded,
        imagesTotal: result.imagesTotal,
        httpErrors: result.httpErrors,
        exceptions: result.exceptions,
        firstError: result.firstError,
        proxyUsed: result.proxyUsed,
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to generate presentation');
    } finally {
      setGenerating(false);
      setImgProgress(null);
      setBuildingPpt(false);
    }
  };

  const reset = () => {
    setFile(null);
    setPreview(null);
    setError(null);
    setDone(null);
    setImgProgress(null);
    if (inputRef.current) inputRef.current.value = '';
  };

  const progressLabel = () => {
    if (buildingPpt) return 'Building presentation…';
    if (imgProgress) {
      const { loaded, total } = imgProgress;
      if (total === 0) return 'Building presentation…';
      return `Fetching images: ${loaded} / ${total}`;
    }
    return 'Preparing…';
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
        <Image src="/perigee-logo.png" alt="Perigee" width={44} height={44} style={{ objectFit: 'contain' }} />
      </header>

      <main style={{ maxWidth: '680px', margin: '0 auto', padding: '2.5rem 1rem' }}>
        <h1 style={{ fontSize: '1.5rem', fontWeight: 700, color: '#242424', marginBottom: '0.25rem' }}>
          Visual Merch Presentation Generator
        </h1>
        <p style={{ color: '#6b7280', marginBottom: '2rem', fontSize: '0.95rem' }}>
          Upload your survey export and download a fully branded PowerPoint in one click.
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

        {/* ── Step 1: Upload ── */}
        <div style={card}>
          <StepLabel n={1} label="Upload Survey Excel" />
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
          <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
            onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
        </div>

        {/* ── Step 2: Preview ── */}
        {preview && (
          <div style={card}>
            <StepLabel n={2} label="Preview" />
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.75rem', marginBottom: '1rem' }}>
              <StatCard label="Total Surveys" value={preview.totalRows} />
              <StatCard label="Unique Users" value={preview.uniqueUsers.length} />
              <StatCard label="Unique Stores" value={preview.uniqueStores.length} />
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

        {/* ── Proxy config (shown when images are present) ── */}
        {preview && preview.totalImages > 0 && (
          <div style={{ ...card, borderColor: proxyUrl ? '#76bd22' : '#f59e0b', background: proxyUrl ? '#fff' : '#fffbeb' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ fontSize: '1rem' }}>{proxyUrl ? '✅' : '⚠️'}</span>
                <div>
                  <div style={{ fontWeight: 700, color: '#242424', fontSize: '0.9rem' }}>
                    {proxyUrl ? 'Image proxy configured' : 'Image proxy required'}
                  </div>
                  <div style={{ color: '#6b7280', fontSize: '0.78rem', marginTop: '0.1rem' }}>
                    {proxyUrl
                      ? proxyUrl
                      : `${preview.totalImages} images need a Cloudflare Worker proxy to embed correctly.`}
                  </div>
                </div>
              </div>
              <button
                onClick={() => setProxyOpen(o => !o)}
                style={{ background: 'none', border: '1px solid #d1d5db', borderRadius: 6, padding: '0.3rem 0.75rem', cursor: 'pointer', color: '#6b7280', fontSize: '0.8rem', flexShrink: 0 }}
              >
                {proxyOpen ? 'Cancel' : proxyUrl ? 'Change' : 'Set up'}
              </button>
            </div>

            {proxyOpen && (
              <div style={{ marginTop: '1rem', paddingTop: '1rem', borderTop: '1px solid #e5e7eb' }}>
                <p style={{ fontSize: '0.82rem', color: '#374151', marginBottom: '0.75rem', lineHeight: 1.5 }}>
                  <strong>One-time setup (5 min):</strong> Create a free{' '}
                  <a href="https://workers.cloudflare.com" target="_blank" rel="noreferrer" style={{ color: '#76bd22' }}>Cloudflare Workers</a>{' '}
                  account → <em>Workers &amp; Pages → Create Worker</em> → paste the script below → Deploy.
                  Copy the <code style={{ background: '#f3f4f6', padding: '0 4px', borderRadius: 3 }}>workers.dev</code> URL and paste it here.
                </p>
                <pre style={{
                  background: '#1e1e1e', color: '#d4d4d4', borderRadius: 8, padding: '0.75rem 1rem',
                  fontSize: '0.72rem', overflowX: 'auto', marginBottom: '0.75rem', lineHeight: 1.6,
                }}>
{`export default {
  async fetch(request) {
    const url = new URL(request.url);
    const target = url.searchParams.get('url');
    if (!target?.startsWith('https://live.perigeeportal.co.za/'))
      return new Response('Bad request', { status: 400 });
    const res = await fetch(target, {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Referer': 'https://live.perigeeportal.co.za/' }
    });
    const headers = new Headers(res.headers);
    headers.set('Access-Control-Allow-Origin', '*');
    return new Response(res.body, { status: res.status, headers });
  }
};`}
                </pre>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                  <input
                    type="url"
                    placeholder="https://perigee-proxy.yourname.workers.dev"
                    value={proxyInput}
                    onChange={e => setProxyInput(e.target.value)}
                    style={{ ...inputStyle, flex: 1 }}
                  />
                  <button
                    onClick={saveProxy}
                    disabled={!proxyInput.trim()}
                    style={{
                      background: proxyInput.trim() ? '#76bd22' : '#a8d97b',
                      color: '#fff', border: 'none', borderRadius: 8,
                      padding: '0 1.25rem', cursor: proxyInput.trim() ? 'pointer' : 'not-allowed',
                      fontWeight: 700, fontSize: '0.9rem', flexShrink: 0,
                    }}
                  >
                    Save
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── Step 3: Generate ── */}
        {preview && (
          <div style={card}>
            <StepLabel n={3} label="Generate Presentation" />

            {done ? (
              <div style={{ background: '#f0fce4', border: '1px solid #76bd22', borderRadius: 8, padding: '1.25rem', textAlign: 'center' }}>
                <div style={{ fontSize: '1.75rem', marginBottom: '0.25rem' }}>🎉</div>
                <div style={{ fontWeight: 700, color: '#242424', marginBottom: '0.35rem' }}>Presentation downloaded!</div>
                <div style={{ color: '#6b7280', fontSize: '0.85rem', marginBottom: '0.5rem' }}>
                  Check your Downloads folder for <strong>FairPrice_VisualMerch.pptx</strong>
                </div>
                {done.imagesTotal > 0 && (
                  <div style={{ marginBottom: '1rem', fontSize: '0.82rem' }}>
                    {done.imagesLoaded === done.imagesTotal ? (
                      <div style={{ color: '#5e9a18' }}>All {done.imagesTotal} images embedded successfully.</div>
                    ) : (
                      <div>
                        <div style={{ color: '#b45309', marginBottom: '0.35rem' }}>
                          {done.imagesLoaded} of {done.imagesTotal} images embedded.
                        </div>
                        <div style={{ background: '#fef3c7', border: '1px solid #fbbf24', borderRadius: 6, padding: '0.5rem 0.75rem', color: '#92400e', textAlign: 'left', lineHeight: 1.5 }}>
                          <div><strong>Proxy used:</strong> <code style={{ fontSize: '0.75rem' }}>{done.proxyUsed}</code></div>
                          {done.httpErrors > 0 && <div><strong>HTTP errors:</strong> {done.httpErrors} (proxy reached Perigee but got blocked — try a different proxy)</div>}
                          {done.exceptions > 0 && <div><strong>Network errors:</strong> {done.exceptions} (proxy unreachable or CORS error)</div>}
                          {done.firstError && <div><strong>First error:</strong> {done.firstError}</div>}
                        </div>
                      </div>
                    )}
                  </div>
                )}
                <button onClick={reset} style={{ background: 'transparent', border: '1.5px solid #76bd22', color: '#5e9a18', borderRadius: 6, padding: '0.4rem 1rem', cursor: 'pointer', fontWeight: 600, fontSize: '0.875rem' }}>
                  Build another
                </button>
              </div>
            ) : (
              <>
                {!generating && (
                  <p style={{ color: '#6b7280', fontSize: '0.875rem', marginBottom: '1.25rem' }}>
                    Images are fetched directly from Perigee in your browser, then the presentation is built locally.
                    {preview.totalImages > 0 && ` Fetching ${preview.totalImages} images — this may take a moment.`}
                  </p>
                )}

                {generating && imgProgress && (
                  <div style={{ marginBottom: '1.25rem' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.4rem', fontSize: '0.85rem', color: '#242424' }}>
                      <span>{progressLabel()}</span>
                      {imgProgress.total > 0 && !buildingPpt && (
                        <span style={{ color: '#76bd22', fontWeight: 600 }}>
                          {imgProgress.loaded} / {imgProgress.total}
                        </span>
                      )}
                    </div>
                    <div style={{ height: 8, background: '#e5e7eb', borderRadius: 4, overflow: 'hidden' }}>
                      <div style={{
                        height: '100%', borderRadius: 4, background: '#76bd22', transition: 'width 0.3s',
                        width: imgProgress.total > 0
                          ? `${Math.round((imgProgress.loaded / imgProgress.total) * 100)}%`
                          : buildingPpt ? '100%' : '10%',
                      }} />
                    </div>
                  </div>
                )}

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
                  {generating ? <><Spinner color="#fff" /> {progressLabel()}</> : <>▶ Generate Presentation</>}
                </button>
              </>
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

// ── Shared styles ─────────────────────────────────────────────────────────────

const card: React.CSSProperties = {
  background: '#fff', borderRadius: 12, padding: '1.5rem', marginBottom: '1.25rem',
  boxShadow: '0 1px 4px rgba(0,0,0,0.07)', border: '1px solid #e2e2e2',
};

const inputStyle: React.CSSProperties = {
  border: '1.5px solid #d1d5db', borderRadius: 8, padding: '0.65rem 0.875rem',
  fontSize: '0.925rem', color: '#242424', outline: 'none', width: '100%',
};

// ── Sub-components ────────────────────────────────────────────────────────────

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
