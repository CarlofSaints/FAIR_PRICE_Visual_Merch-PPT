import PptxGenJS from 'pptxgenjs';
import type { ParsedData, SurveyRow, UserSummary } from '@/types/survey';
import path from 'path';
import fs from 'fs';

// ─── Brand tokens ────────────────────────────────────────────────────────────
const GREEN = '76bd22';
const RED = 'FF4539';
const DARK = '242424';
const WHITE = 'FFFFFF';
const LIGHT_GRAY = 'F6F6F6';

// Slide dimensions: 10" × 5.625" (widescreen)
const W = 10;
const H = 5.625;

// Header bar
const BAR_H = 0.95;
const BAR_Y = 0;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function logoPath(filename: string): string {
  return path.join(process.cwd(), 'public', filename);
}

function logoBase64(filename: string): string | null {
  try {
    const p = logoPath(filename);
    if (!fs.existsSync(p)) return null;
    return fs.readFileSync(p).toString('base64');
  } catch {
    return null;
  }
}

function addGreenBar(slide: PptxGenJS.Slide, title: string, subtitle?: string, dateStr?: string) {
  slide.addShape('rect', { x: 0, y: BAR_Y, w: W, h: BAR_H, fill: { color: GREEN }, line: { color: GREEN } });

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
      x: 7.5, y: BAR_Y, w: 2.25, h: BAR_H,
      fontSize: 11, color: WHITE, fontFace: 'Calibri', align: 'right', valign: 'middle',
    });
  }
}

function addLogos(slide: PptxGenJS.Slide, opts: { fpLeft?: boolean; pgRight?: boolean; small?: boolean } = {}) {
  const fpB64 = logoBase64('fairprice-logo.png');
  const pgB64 = logoBase64('perigee-logo.png');

  if (opts.fpLeft && fpB64) {
    slide.addImage({ data: 'image/png;base64,' + fpB64, x: 0.25, y: 0.12, w: 1.4, h: 0.65 });
  }
  if (opts.pgRight && pgB64) {
    const sz = opts.small ? 0.55 : 0.75;
    slide.addImage({ data: 'image/png;base64,' + pgB64, x: W - sz - 0.2, y: opts.small ? BAR_H - sz - 0.05 : H - sz - 0.15, w: sz, h: sz });
  }
}

// ─── Slide 1: Title ──────────────────────────────────────────────────────────
function addTitleSlide(pptx: PptxGenJS, data: ParsedData) {
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: GREEN }, line: { color: GREEN } });

  const fpB64 = logoBase64('fairprice-logo.png');
  if (fpB64) {
    slide.addImage({ data: 'image/png;base64,' + fpB64, x: 0.4, y: 0.3, w: 2.2, h: 1.0 });
  }

  const pgB64 = logoBase64('perigee-logo.png');
  if (pgB64) {
    slide.addImage({ data: 'image/png;base64,' + pgB64, x: W - 1.3, y: H - 1.3, w: 1.1, h: 1.1 });
  }

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

// ─── Slide 2: Summary ────────────────────────────────────────────────────────
function addSummarySlide(pptx: PptxGenJS, data: ParsedData, summaries: UserSummary[]) {
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF' } });

  // Green bar
  slide.addShape('rect', { x: 0, y: 0, w: W, h: BAR_H, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addText('Survey Summary', {
    x: 0.25, y: 0, w: W - 0.5, h: BAR_H,
    fontSize: 18, bold: true, color: WHITE, fontFace: 'Calibri', valign: 'middle',
  });

  const fpB64 = logoBase64('fairprice-logo.png');
  if (fpB64) {
    slide.addImage({ data: 'image/png;base64,' + fpB64, x: W - 1.6, y: 0.12, w: 1.35, h: 0.65 });
  }

  // Stats row
  slide.addText(
    `Total Users: ${data.uniqueUsers.length}    |    Total Unique Stores: ${data.uniqueStores.length}    |    Total Surveys: ${data.totalRows}`,
    { x: 0.3, y: 1.05, w: W - 0.6, h: 0.4, fontSize: 12, color: DARK, fontFace: 'Calibri', bold: true }
  );

  // Table header
  const days = data.uniqueDays;
  const colNames = ['Name', 'Unique Stores', 'Total Surveys', ...days];
  const colW = [2.2, 1.4, 1.4, ...days.map(() => Math.max(0.9, (W - 5.0) / Math.max(days.length, 1)))];

  const headerRow = colNames.map((n, i) => ({
    text: n,
    options: { bold: true, color: WHITE, fill: GREEN, fontSize: 10, align: 'center' as const, fontFace: 'Calibri' },
  }));

  const tableRows = [headerRow, ...summaries.map(s => [
    { text: s.fullName, options: { fontSize: 10, fontFace: 'Calibri', color: DARK } },
    { text: String(s.uniqueStores), options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK } },
    { text: String(s.totalSurveys), options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK } },
    ...days.map(d => ({
      text: String(s.dayCounts[d] || 0),
      options: { fontSize: 10, fontFace: 'Calibri', align: 'center' as const, color: DARK },
    })),
  ])];

  slide.addTable(tableRows, {
    x: 0.3, y: 1.55, w: W - 0.6,
    colW,
    border: { type: 'solid', color: 'CCCCCC', pt: 0.5 },
    rowH: 0.35,
  });
}

// ─── Store slides ─────────────────────────────────────────────────────────────
async function addStoreSlides(pptx: PptxGenJS, row: SurveyRow, imageCache: Map<string, string>) {
  // ── Summary slide ───────────────────────────────────────────────────────────
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF' } });
  addGreenBar(slide, row.store, row.fullName, row.date);

  const BODY_TOP = BAR_H + 0.15;
  const BODY_H = H - BODY_TOP - 0.1;

  // Collect non-empty sections
  const activeSections = row.sections.filter(s => s.qaPairs.length > 0);

  if (activeSections.length === 0) {
    slide.addText('No survey data recorded.', {
      x: 0.3, y: BODY_TOP + 0.3, w: W - 0.6, h: 0.5,
      fontSize: 11, color: '888888', fontFace: 'Calibri', italic: true,
    });
  } else {
    // Two-column layout for sections
    const allPairs: Array<{ section: string; q: string; a: string }> = [];
    for (const sec of activeSections) {
      for (const qa of sec.qaPairs) {
        allPairs.push({ section: sec.name, q: qa.question, a: qa.answer });
      }
    }

    // Split into chunks of ~22 pairs per slide
    const CHUNK = 22;
    const chunks: typeof allPairs[] = [];
    for (let i = 0; i < allPairs.length; i += CHUNK) chunks.push(allPairs.slice(i, i + CHUNK));

    for (let ci = 0; ci < chunks.length; ci++) {
      const targetSlide = ci === 0 ? slide : (() => {
        const s2 = pptx.addSlide();
        s2.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF' } });
        addGreenBar(s2, row.store, row.fullName + (ci > 0 ? ' (cont.)' : ''), row.date);
        return s2;
      })();

      const chunk = chunks[ci];
      const half = Math.ceil(chunk.length / 2);
      const left = chunk.slice(0, half);
      const right = chunk.slice(half);

      const renderCol = (pairs: typeof allPairs, xStart: number, colWidth: number) => {
        let y = BODY_TOP;
        let lastSection = '';
        for (const p of pairs) {
          if (p.section !== lastSection) {
            lastSection = p.section;
            targetSlide.addText(p.section.toUpperCase(), {
              x: xStart, y, w: colWidth, h: 0.22,
              fontSize: 8, bold: true, color: GREEN, fontFace: 'Calibri',
            });
            y += 0.22;
          }
          const qText = p.q.length > 55 ? p.q.slice(0, 52) + '…' : p.q;
          const lineH = 0.22;
          targetSlide.addText([
            { text: qText + '  ', options: { color: '555555', fontSize: 8 } },
            { text: p.a, options: { color: DARK, bold: true, fontSize: 8 } },
          ], { x: xStart, y, w: colWidth, h: lineH, fontFace: 'Calibri', valign: 'top' });
          y += lineH;
          if (y > H - 0.2) break;
        }
      };

      renderCol(left, 0.25, 4.6);
      if (right.length > 0) renderCol(right, 5.15, 4.6);
    }
  }

  // ── Image slides ────────────────────────────────────────────────────────────
  for (const img of row.imageEntries) {
    const imgB64 = imageCache.get(img.imageUrl);
    const imgSlide = pptx.addSlide();
    imgSlide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF' } });
    addGreenBar(imgSlide, row.store, row.fullName, row.date);

    // Section label
    imgSlide.addText(img.sectionName.toUpperCase(), {
      x: 0.25, y: BAR_H + 0.1, w: 4, h: 0.25,
      fontSize: 8, bold: true, color: GREEN, fontFace: 'Calibri',
    });

    const IMG_TOP = BAR_H + 0.4;
    const IMG_H = 3.2;
    const IMG_W = 5.5;
    const IMG_X = (W - IMG_W) / 2;

    if (imgB64) {
      imgSlide.addImage({
        data: 'image/jpeg;base64,' + imgB64,
        x: IMG_X, y: IMG_TOP, w: IMG_W, h: IMG_H,
        sizing: { type: 'contain', w: IMG_W, h: IMG_H },
      });
    } else {
      imgSlide.addShape('rect', {
        x: IMG_X, y: IMG_TOP, w: IMG_W, h: IMG_H,
        fill: { color: LIGHT_GRAY }, line: { color: 'CCCCCC' },
      });
      imgSlide.addText('Image unavailable', {
        x: IMG_X, y: IMG_TOP + IMG_H / 2 - 0.2, w: IMG_W, h: 0.4,
        fontSize: 11, color: '999999', fontFace: 'Calibri', align: 'center',
      });
    }

    const captionTop = IMG_TOP + IMG_H + 0.1;
    if (img.precedingQuestion) {
      imgSlide.addText(`Q: ${img.precedingQuestion}`, {
        x: 0.25, y: captionTop, w: W - 0.5, h: 0.25,
        fontSize: 9, color: '555555', fontFace: 'Calibri', italic: true,
      });
      if (img.precedingAnswer) {
        imgSlide.addText(`A: ${img.precedingAnswer}`, {
          x: 0.25, y: captionTop + 0.24, w: W - 0.5, h: 0.25,
          fontSize: 9, bold: true, color: DARK, fontFace: 'Calibri',
        });
      }
    }
    if (img.imageHeader) {
      imgSlide.addText(img.imageHeader, {
        x: 0.25, y: captionTop + 0.5, w: W - 0.5, h: 0.22,
        fontSize: 8, color: '888888', fontFace: 'Calibri', italic: true,
      });
    }
  }
}

// ─── Final slide ─────────────────────────────────────────────────────────────
function addThankYouSlide(pptx: PptxGenJS) {
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: GREEN }, line: { color: GREEN } });

  slide.addText('Thank You', {
    x: 0.5, y: 1.5, w: W - 1, h: 1.5,
    fontSize: 48, bold: true, color: WHITE, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  const fpB64 = logoBase64('fairprice-logo.png');
  if (fpB64) {
    slide.addImage({ data: 'image/png;base64,' + fpB64, x: 1.5, y: 3.5, w: 2.5, h: 1.1 });
  }

  const pgB64 = logoBase64('perigee-logo.png');
  if (pgB64) {
    slide.addImage({ data: 'image/png;base64,' + pgB64, x: W - 4.0, y: 3.4, w: 1.2, h: 1.2 });
  }
}

// ─── Image pre-fetcher ───────────────────────────────────────────────────────
async function fetchImages(rows: SurveyRow[]): Promise<Map<string, string>> {
  const allUrls = new Set<string>();
  for (const row of rows) {
    for (const img of row.imageEntries) {
      if (img.imageUrl) allUrls.add(img.imageUrl);
    }
  }

  const cache = new Map<string, string>();
  const results = await Promise.allSettled(
    [...allUrls].map(async url => {
      const res = await fetch(url, { signal: AbortSignal.timeout(15000) });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buf = await res.arrayBuffer();
      cache.set(url, Buffer.from(buf).toString('base64'));
    })
  );

  let failed = 0;
  results.forEach((r, i) => {
    if (r.status === 'rejected') {
      failed++;
      console.warn('Image fetch failed:', [...allUrls][i], r.reason?.message);
    }
  });
  if (failed > 0) console.warn(`${failed}/${allUrls.size} images failed to load`);

  return cache;
}

// ─── Main export ─────────────────────────────────────────────────────────────
export async function buildPptx(data: ParsedData, summaries: UserSummary[]): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 10" × 5.625"
  pptx.author = 'Perigee Field Goose';
  pptx.subject = 'Fair Price Visual Merch Report';

  // Pre-fetch all images
  const imageCache = await fetchImages(data.rows);

  addTitleSlide(pptx, data);
  addSummarySlide(pptx, data, summaries);

  for (const row of data.rows) {
    await addStoreSlides(pptx, row, imageCache);
  }

  addThankYouSlide(pptx);

  const result = await pptx.write({ outputType: 'nodebuffer' });
  return result as Buffer;
}
