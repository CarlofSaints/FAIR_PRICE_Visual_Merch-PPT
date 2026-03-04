import PptxGenJS from 'pptxgenjs';
import type { UserSummary } from '@/types/survey';

// ─── Brand tokens ─────────────────────────────────────────────────────────────
const GREEN = '76bd22';
const DARK = '242424';
const WHITE = 'FFFFFF';
const LIGHT_GRAY = 'F6F6F6';

const W = 10;
const H = 5.625;
const BAR_H = 0.95;
const LINE_H = 0.28;
const SEC_H = 0.28;
const CHUNK = 18;

// ─── Serialised types (Date → string after JSON roundtrip) ────────────────────
export interface ImageEntryJSON {
  sectionName: string;
  imageUrl: string;
  imageHeader: string;
  precedingQuestion: string;
  precedingAnswer: string;
}

export interface QAPairJSON {
  question: string;
  answer: string;
  imageUrl?: string;
}

export interface SurveySectionJSON {
  name: string;
  qaPairs: QAPairJSON[];
}

export interface SurveyRowJSON {
  id: string | number;
  email: string;
  firstName: string;
  lastName: string;
  fullName: string;
  store: string;
  storeCode: string;
  province: string;
  date: string;
  dateRaw: string | null;
  dayOfWeek: string;
  sections: SurveySectionJSON[];
  imageEntries: ImageEntryJSON[];
}

export interface ParsedDataJSON {
  rows: SurveyRowJSON[];
  uniqueUsers: string[];
  uniqueStores: string[];
  uniqueDays: string[];
  dateRange: { from: string; to: string };
  totalRows: number;
}

export type ProgressCallback = (loaded: number, total: number) => void;

// Proxy base URL — set to a Cloudflare Worker URL to bypass Perigee CORS.
// Falls back to the same-origin Vercel route (which is IP-blocked by Perigee).
let _proxyBase = '/api/proxy-image';
export function setProxyBase(url: string) {
  _proxyBase = url.replace(/\/$/, '');
}

// ─── Image helpers ────────────────────────────────────────────────────────────

async function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function fetchAsBase64(url: string): Promise<string | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return blobToBase64(await res.blob());
  } catch {
    return null;
  }
}

function compressViaCanvas(b64: string, mime: string): Promise<string> {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const MAX = 1400;
      let w = img.naturalWidth;
      let h = img.naturalHeight;
      if (w > MAX || h > MAX) {
        const scale = Math.min(MAX / w, MAX / h);
        w = Math.round(w * scale);
        h = Math.round(h * scale);
      }
      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      canvas.getContext('2d')!.drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.82).split(',')[1]);
    };
    img.onerror = () => resolve(b64);
    img.src = `data:${mime};base64,${b64}`;
  });
}

type FetchOutcome =
  | { ok: true; b64: string; mime: string }
  | { ok: false; reason: 'http'; status: number }
  | { ok: false; reason: 'exception'; msg: string };

async function fetchImageDiag(url: string): Promise<FetchOutcome> {
  const proxyUrl = `${_proxyBase}?url=${encodeURIComponent(url)}`;
  try {
    const res = await fetch(proxyUrl);
    if (!res.ok) return { ok: false, reason: 'http', status: res.status };
    const blob = await res.blob();
    const mime = blob.type || 'image/jpeg';
    const rawB64 = await blobToBase64(blob);
    const compressed = await compressViaCanvas(rawB64, mime);
    return { ok: true, b64: compressed, mime: 'image/jpeg' };
  } catch (e) {
    return { ok: false, reason: 'exception', msg: String(e) };
  }
}

// ─── Main export ──────────────────────────────────────────────────────────────

export interface BuildResult {
  blob: Blob;
  imagesEmbedded: number;  // images fetched + embedded as data
  imagesLinked: number;    // images not fetched — clickable links in PPT instead
  imagesTotal: number;
  httpErrors: number;
  exceptions: number;
  firstError: string;
  proxyUsed: string;
}

export async function buildPptxBrowser(
  data: ParsedDataJSON,
  summaries: UserSummary[],
  onProgress: ProgressCallback,
): Promise<BuildResult> {
  // Load logos from same-origin public folder (no CORS)
  const [fpB64, pgB64] = await Promise.all([
    fetchAsBase64('/fairprice-logo.png'),
    fetchAsBase64('/perigee-logo.png'),
  ]);

  // Collect all unique image URLs
  const allUrls = new Set<string>();
  for (const row of data.rows) {
    for (const img of row.imageEntries) {
      if (img.imageUrl) allUrls.add(img.imageUrl);
    }
  }
  const urlList = [...allUrls];
  const imageCache = new Map<string, { b64: string; mime: string }>();
  let httpErrors = 0;
  let exceptions = 0;
  let firstError = '';

  // Fetch + compress images with progress
  let loaded = 0;
  onProgress(0, urlList.length);
  await Promise.allSettled(
    urlList.map(async url => {
      const outcome = await fetchImageDiag(url);
      if (outcome.ok) {
        imageCache.set(url, { b64: outcome.b64, mime: outcome.mime });
      } else if (outcome.reason === 'http') {
        httpErrors++;
        if (!firstError) firstError = `HTTP ${outcome.status} from proxy`;
      } else {
        exceptions++;
        if (!firstError) firstError = outcome.msg;
      }
      loaded++;
      onProgress(loaded, urlList.length);
    }),
  );

  // ── Build Presentation ──────────────────────────────────────────────────────
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Perigee Field Goose';
  pptx.subject = 'Fair Price Visual Merch Report';

  type Slide = ReturnType<typeof pptx.addSlide>;

  function addGreenBar(slide: Slide, title: string, subtitle?: string, dateStr?: string) {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: BAR_H, fill: { color: GREEN }, line: { type: 'none' } });
    slide.addText(title, {
      x: 0.25, y: 0.07, w: dateStr ? 7 : 9.5, h: 0.5,
      fontSize: 16, bold: true, color: WHITE, fontFace: 'Calibri', valign: 'middle',
    });
    if (subtitle) {
      slide.addText(subtitle, {
        x: 0.25, y: 0.52, w: dateStr ? 7 : 9.5, h: 0.38,
        fontSize: 11, color: WHITE, fontFace: 'Calibri', valign: 'middle',
      });
    }
    if (dateStr) {
      slide.addText(dateStr, {
        x: 7.5, y: 0, w: 2.25, h: BAR_H,
        fontSize: 11, color: WHITE, fontFace: 'Calibri', align: 'right', valign: 'middle',
      });
    }
  }

  function renderCol(
    slide: Slide,
    pairs: Array<{ section: string; q: string; a: string }>,
    xStart: number,
    colWidth: number,
    bodyTop: number,
  ) {
    let y = bodyTop;
    let lastSection = '';
    for (const p of pairs) {
      if (p.section !== lastSection) {
        lastSection = p.section;
        slide.addText(p.section.toUpperCase(), {
          x: xStart, y, w: colWidth, h: SEC_H,
          fontSize: 10, bold: true, color: GREEN, fontFace: 'Calibri',
        });
        y += SEC_H;
      }
      const qText = p.q.length > 60 ? p.q.slice(0, 57) + '…' : p.q;
      slide.addText(
        [
          { text: qText + '  ', options: { color: '555555', fontSize: 10 } },
          { text: p.a, options: { color: DARK, bold: true, fontSize: 10 } },
        ],
        { x: xStart, y, w: colWidth, h: LINE_H, fontFace: 'Calibri', valign: 'top' },
      );
      y += LINE_H;
      if (y > H - 0.2) break;
    }
  }

  // ── Title slide ─────────────────────────────────────────────────────────────
  {
    const slide = pptx.addSlide();
    slide.background = { color: GREEN };
    if (fpB64) slide.addImage({ data: 'image/png;base64,' + fpB64, x: 0.4, y: 0.3, w: 2.2, h: 1.0 });
    if (pgB64) slide.addImage({ data: 'image/png;base64,' + pgB64, x: W - 1.3, y: H - 1.3, w: 1.1, h: 1.1 });
    slide.addText('Visual Merch', {
      x: 0.5, y: 1.7, w: W - 1, h: 1.2,
      fontSize: 44, bold: true, color: WHITE, fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    const { from, to } = data.dateRange;
    const range = from === to || !to ? from : `${from} – ${to}`;
    if (range) {
      slide.addText(range, {
        x: 0.5, y: 3.0, w: W - 1, h: 0.6,
        fontSize: 20, color: WHITE, fontFace: 'Calibri', align: 'center', valign: 'middle',
      });
    }
  }

  // ── Summary slide ────────────────────────────────────────────────────────────
  {
    const slide = pptx.addSlide();
    slide.addShape('rect', { x: 0, y: 0, w: W, h: BAR_H, fill: { color: GREEN }, line: { type: 'none' } });
    slide.addText('Survey Summary', {
      x: 0.25, y: 0, w: W - 0.5, h: BAR_H,
      fontSize: 18, bold: true, color: WHITE, fontFace: 'Calibri', valign: 'middle',
    });
    if (fpB64) slide.addImage({ data: 'image/png;base64,' + fpB64, x: W - 1.6, y: 0.12, w: 1.35, h: 0.65 });
    slide.addText(
      `Total Users: ${data.uniqueUsers.length}    |    Total Unique Stores: ${data.uniqueStores.length}    |    Total Surveys: ${data.totalRows}`,
      { x: 0.3, y: 1.05, w: W - 0.6, h: 0.4, fontSize: 12, color: DARK, fontFace: 'Calibri', bold: true },
    );
    const days = data.uniqueDays;
    const colW = [2.2, 1.4, 1.4, ...days.map(() => Math.max(0.9, (W - 5.0) / Math.max(days.length, 1)))];
    const headerRow = ['Name', 'Unique Stores', 'Total Surveys', ...days].map(n => ({
      text: n,
      options: { bold: true, color: WHITE, fill: GREEN, fontSize: 10, align: 'center' as const, fontFace: 'Calibri' },
    }));
    const tableRows = [
      headerRow,
      ...summaries.map(s => [
        { text: s.fullName, options: { fontSize: 10, fontFace: 'Calibri', color: DARK } },
        { text: String(s.uniqueStores), options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK } },
        { text: String(s.totalSurveys), options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK } },
        ...days.map(d => ({
          text: String(s.dayCounts[d] || 0),
          options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK },
        })),
      ]),
    ];
    slide.addTable(tableRows, {
      x: 0.3, y: 1.55, w: W - 0.6,
      colW,
      border: { type: 'solid', color: 'CCCCCC', pt: 0.5 },
      rowH: 0.35,
    });
  }

  // ── Per-row store slides ─────────────────────────────────────────────────────
  for (const row of data.rows) {
    const BODY_TOP = BAR_H + 0.15;

    // Q&A summary slide
    const slide = pptx.addSlide();
    addGreenBar(slide, row.store, row.fullName, row.date);

    const activeSections = row.sections.filter(s => s.qaPairs.length > 0);

    if (activeSections.length === 0) {
      slide.addText('No survey data recorded.', {
        x: 0.3, y: BODY_TOP + 0.3, w: W - 0.6, h: 0.5,
        fontSize: 11, color: '888888', fontFace: 'Calibri', italic: true,
      });
    } else {
      const allPairs: Array<{ section: string; q: string; a: string }> = [];
      for (const sec of activeSections) {
        for (const qa of sec.qaPairs) {
          allPairs.push({ section: sec.name, q: qa.question, a: qa.answer });
        }
      }
      const chunks: (typeof allPairs)[] = [];
      for (let i = 0; i < allPairs.length; i += CHUNK) chunks.push(allPairs.slice(i, i + CHUNK));

      for (let ci = 0; ci < chunks.length; ci++) {
        const targetSlide = ci === 0 ? slide : (() => {
          const s2 = pptx.addSlide();
          addGreenBar(s2, row.store, row.fullName + ' (cont.)', row.date);
          return s2;
        })();
        const chunk = chunks[ci];
        const half = Math.ceil(chunk.length / 2);
        renderCol(targetSlide, chunk.slice(0, half), 0.25, 4.6, BODY_TOP);
        if (chunk.slice(half).length > 0) renderCol(targetSlide, chunk.slice(half), 5.15, 4.6, BODY_TOP);
      }
    }

    // Image slides
    for (const img of row.imageEntries) {
      const cached = imageCache.get(img.imageUrl);
      const imgSlide = pptx.addSlide();
      addGreenBar(imgSlide, row.store, row.fullName, row.date);

      imgSlide.addText(img.sectionName.toUpperCase(), {
        x: 0.25, y: BAR_H + 0.1, w: 4, h: 0.28,
        fontSize: 10, bold: true, color: GREEN, fontFace: 'Calibri',
      });

      const IMG_TOP = BAR_H + 0.45;
      const IMG_H = 3.1;
      const IMG_W = 5.5;
      const IMG_X = (W - IMG_W) / 2;

      if (cached) {
        imgSlide.addImage({
          data: `${cached.mime};base64,${cached.b64}`,
          x: IMG_X, y: IMG_TOP, w: IMG_W, h: IMG_H,
          sizing: { type: 'contain', w: IMG_W, h: IMG_H },
        });
      } else {
        // Clickable link — opens image in browser when clicked in PowerPoint
        imgSlide.addShape('rect', {
          x: IMG_X, y: IMG_TOP, w: IMG_W, h: IMG_H,
          fill: { color: 'EDF7E0' },
          line: { color: GREEN, pt: 1.5 },
          hyperlink: { url: img.imageUrl, tooltip: 'Click to open image in Perigee' },
        });
        imgSlide.addText('Click to view image  ↗', {
          x: IMG_X, y: IMG_TOP + IMG_H / 2 - 0.35, w: IMG_W, h: 0.55,
          fontSize: 14, bold: true, color: '3d7a0a', fontFace: 'Calibri', align: 'center',
          hyperlink: { url: img.imageUrl },
        });
        imgSlide.addText('(opens in browser)', {
          x: IMG_X, y: IMG_TOP + IMG_H / 2 + 0.25, w: IMG_W, h: 0.35,
          fontSize: 10, color: '6b9e30', fontFace: 'Calibri', align: 'center', italic: true,
        });
      }

      const captionTop = IMG_TOP + IMG_H + 0.1;
      if (img.precedingQuestion) {
        imgSlide.addText(`Q: ${img.precedingQuestion}`, {
          x: 0.25, y: captionTop, w: W - 0.5, h: 0.27,
          fontSize: 10, color: '555555', fontFace: 'Calibri', italic: true,
        });
        if (img.precedingAnswer) {
          imgSlide.addText(`A: ${img.precedingAnswer}`, {
            x: 0.25, y: captionTop + 0.27, w: W - 0.5, h: 0.27,
            fontSize: 10, bold: true, color: DARK, fontFace: 'Calibri',
          });
        }
      }
      if (img.imageHeader) {
        imgSlide.addText(img.imageHeader, {
          x: 0.25, y: captionTop + 0.56, w: W - 0.5, h: 0.25,
          fontSize: 9, color: '888888', fontFace: 'Calibri', italic: true,
        });
      }
    }
  }

  // ── Thank You slide ──────────────────────────────────────────────────────────
  {
    const slide = pptx.addSlide();
    slide.background = { color: GREEN };
    slide.addText('Thank You', {
      x: 0.5, y: 1.5, w: W - 1, h: 1.5,
      fontSize: 48, bold: true, color: WHITE, fontFace: 'Calibri', align: 'center', valign: 'middle',
    });
    if (fpB64) slide.addImage({ data: 'image/png;base64,' + fpB64, x: 1.5, y: 3.5, w: 2.5, h: 1.1 });
    if (pgB64) slide.addImage({ data: 'image/png;base64,' + pgB64, x: W - 4.0, y: 3.4, w: 1.2, h: 1.2 });
  }

  const result = await pptx.write({ outputType: 'blob' });
  return {
    blob: result as Blob,
    imagesEmbedded: imageCache.size,
    imagesLinked: urlList.length - imageCache.size,
    imagesTotal: urlList.length,
    httpErrors,
    exceptions,
    firstError,
    proxyUsed: _proxyBase,
  };
}
