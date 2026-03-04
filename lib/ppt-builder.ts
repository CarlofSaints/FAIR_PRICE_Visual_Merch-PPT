import PptxGenJS from 'pptxgenjs';
import type { ParsedData, SurveyRow, UserSummary } from '@/types/survey';
import path from 'path';
import fs from 'fs';
import sharp from 'sharp';

// ─── Brand tokens ────────────────────────────────────────────────────────────
const GREEN = '76bd22';
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
  slide.addShape('rect', { x: 0, y: BAR_Y, w: W, h: BAR_H, fill: { color: GREEN }, line: { type: 'none' } });

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

// ─── Slide 1: Title ──────────────────────────────────────────────────────────
function addTitleSlide(pptx: PptxGenJS, data: ParsedData) {
  const slide = pptx.addSlide();
  slide.background = { color: GREEN };  // full-bleed background

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
  // No need to paint white rect — slides default to white

  // Green bar
  slide.addShape('rect', { x: 0, y: 0, w: W, h: BAR_H, fill: { color: GREEN }, line: { type: 'none' } });
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

  // Table
  const days = data.uniqueDays;
  const colNames = ['Name', 'Unique Stores', 'Total Surveys', ...days];
  const colW = [2.2, 1.4, 1.4, ...days.map(() => Math.max(0.9, (W - 5.0) / Math.max(days.length, 1)))];

  const headerRow = colNames.map(n => ({
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
async function addStoreSlides(pptx: PptxGenJS, row: SurveyRow, imageCache: Map<string, { b64: string; mime: string }>) {
  // ── Q&A Summary slide(s) ────────────────────────────────────────────────────
  const slide = pptx.addSlide();
  addGreenBar(slide, row.store, row.fullName, row.date);

  const BODY_TOP = BAR_H + 0.15;
  const LINE_H = 0.28;   // increased for 10pt text
  const SEC_H = 0.28;

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

    // ~18 pairs per chunk fits comfortably at 10pt with SEC_H overhead
    const CHUNK = 18;
    const chunks: typeof allPairs[] = [];
    for (let i = 0; i < allPairs.length; i += CHUNK) chunks.push(allPairs.slice(i, i + CHUNK));

    for (let ci = 0; ci < chunks.length; ci++) {
      const targetSlide = ci === 0 ? slide : (() => {
        const s2 = pptx.addSlide();
        addGreenBar(s2, row.store, row.fullName + ' (cont.)', row.date);
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
              x: xStart, y, w: colWidth, h: SEC_H,
              fontSize: 10, bold: true, color: GREEN, fontFace: 'Calibri',
            });
            y += SEC_H;
          }
          const qText = p.q.length > 60 ? p.q.slice(0, 57) + '…' : p.q;
          targetSlide.addText([
            { text: qText + '  ', options: { color: '555555', fontSize: 10 } },
            { text: p.a, options: { color: DARK, bold: true, fontSize: 10 } },
          ], { x: xStart, y, w: colWidth, h: LINE_H, fontFace: 'Calibri', valign: 'top' });
          y += LINE_H;
          if (y > H - 0.2) break;
        }
      };

      renderCol(left, 0.25, 4.6);
      if (right.length > 0) renderCol(right, 5.15, 4.6);
    }
  }

  // ── Image slides ────────────────────────────────────────────────────────────
  for (const img of row.imageEntries) {
    const cached = imageCache.get(img.imageUrl);
    const imgSlide = pptx.addSlide();
    addGreenBar(imgSlide, row.store, row.fullName, row.date);

    // Section label
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

// ─── Final slide ─────────────────────────────────────────────────────────────
function addThankYouSlide(pptx: PptxGenJS) {
  const slide = pptx.addSlide();
  slide.background = { color: GREEN };  // full-bleed background

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
async function fetchImages(rows: SurveyRow[], perigeeCookie: string): Promise<Map<string, { b64: string; mime: string }>> {
  const allUrls = new Set<string>();
  for (const row of rows) {
    for (const img of row.imageEntries) {
      if (img.imageUrl) allUrls.add(img.imageUrl);
    }
  }

  const cache = new Map<string, { b64: string; mime: string }>();
  const urlList = [...allUrls];

  console.log(`Fetching ${urlList.length} unique images…`);

  const results = await Promise.allSettled(
    urlList.map(async url => {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 20000);
      try {
        const res = await fetch(url, {
          signal: controller.signal,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/jpeg,image/png,image/webp,image/*,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://live.perigeeportal.co.za/',
            'Cookie': perigeeCookie,
          },
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const rawBuf = Buffer.from(await res.arrayBuffer());

        // Compress: resize to max 1400px wide, JPEG 82% quality
        const compressed = await sharp(rawBuf)
          .resize({ width: 1400, height: 1400, fit: 'inside', withoutEnlargement: true })
          .jpeg({ quality: 82, progressive: true })
          .toBuffer();

        cache.set(url, { b64: compressed.toString('base64'), mime: 'image/jpeg' });
      } finally {
        clearTimeout(timer);
      }
    })
  );

  let failed = 0;
  results.forEach((r, i) => {
    if (r.status === 'rejected') {
      failed++;
      console.error(`Image failed [${i}]: ${urlList[i]} — ${r.reason?.message ?? r.reason}`);
    }
  });
  console.log(`Images loaded: ${urlList.length - failed}/${urlList.length}`);

  return cache;
}

// ─── Main export ─────────────────────────────────────────────────────────────
export async function buildPptx(data: ParsedData, summaries: UserSummary[], perigeeCookie: string): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 10" × 5.625"
  pptx.author = 'Perigee Field Goose';
  pptx.subject = 'Fair Price Visual Merch Report';

  const imageCache = await fetchImages(data.rows, perigeeCookie);

  addTitleSlide(pptx, data);
  addSummarySlide(pptx, data, summaries);

  for (const row of data.rows) {
    await addStoreSlides(pptx, row, imageCache);
  }

  addThankYouSlide(pptx);

  const result = await pptx.write({ outputType: 'nodebuffer' });
  return result as Buffer;
}
