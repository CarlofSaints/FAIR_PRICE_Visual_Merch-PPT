import * as XLSX from 'xlsx';
import type { ParsedData, SurveyRow, SurveySection, QAPair, ImageEntry, UserSummary } from '@/types/survey';

// Fixed metadata column indices (0-based)
const COL = {
  EMAIL: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  STORE: 6,
  STORE_CODE: 7,
  PROVINCE: 8,
  DATE: 9,
  DAY: 15,
  QUESTIONS_START: 16,
};

// Product section boundaries (start index, inclusive)
const SECTIONS: Array<{ name: string; startIdx: number }> = [
  { name: 'Base Sets', startIdx: 16 },
  { name: 'Bedroom Suites', startIdx: 37 },
  { name: 'Lounge Suites', startIdx: 67 },
  { name: 'Coffee Tables', startIdx: 92 },
  { name: 'Kitchen Units', startIdx: 116 },
  { name: 'Fridges & Freezers', startIdx: 145 },
  { name: 'Stoves', startIdx: 175 },
  { name: 'Washing Machines', startIdx: 199 },
];

function getSectionName(colIdx: number): string {
  let name = 'General';
  for (const s of SECTIONS) {
    if (colIdx >= s.startIdx) name = s.name;
    else break;
  }
  return name;
}

function isImageHeader(header: string): boolean {
  const h = header.toLowerCase();
  return h.includes('upload a pic') || h.includes('take a general pic') || h.includes('pic of where');
}

function isImageUrl(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith('https://live.perigeeportal.co.za');
}

function formatDate(raw: unknown): { formatted: string; date: Date | null } {
  if (!raw) return { formatted: '', date: null };
  // XLSX may return a JS Date, a serial number, or a string
  if (raw instanceof Date) {
    return { formatted: formatDateObj(raw), date: raw };
  }
  if (typeof raw === 'number') {
    const d = XLSX.SSF.parse_date_code(raw);
    if (d) {
      const date = new Date(d.y, d.m - 1, d.d);
      return { formatted: formatDateObj(date), date };
    }
  }
  if (typeof raw === 'string') {
    // Perigee uses DD/MM/YYYY — parse manually to avoid JS treating it as MM/DD/YYYY
    const dmyMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmyMatch) {
      const d = new Date(Number(dmyMatch[3]), Number(dmyMatch[2]) - 1, Number(dmyMatch[1]));
      if (!isNaN(d.getTime())) return { formatted: formatDateObj(d), date: d };
    }
    const d = new Date(raw);
    if (!isNaN(d.getTime())) return { formatted: formatDateObj(d), date: d };
    return { formatted: raw, date: null };
  }
  return { formatted: String(raw), date: null };
}

function formatDateObj(d: Date): string {
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

export function parseExcel(buffer: ArrayBuffer): ParsedData {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  if (raw.length < 2) {
    return { rows: [], uniqueUsers: [], uniqueStores: [], uniqueDays: [], dateRange: { from: '', to: '' }, totalRows: 0 };
  }

  const headers = (raw[0] as unknown[]).map(h => (h != null ? String(h) : ''));
  const dataRows = raw.slice(1).filter(r => r.some(v => v != null));

  const rows: SurveyRow[] = dataRows.map(row => {
    const get = (i: number) => (row[i] != null ? String(row[i]).trim() : '');

    const { formatted: dateStr, date: dateObj } = formatDate(row[COL.DATE]);

    // Build sections
    const sectionMap: Map<string, QAPair[]> = new Map();
    const imageEntries: ImageEntry[] = [];

    for (let i = COL.QUESTIONS_START; i < headers.length; i++) {
      const header = headers[i];
      const value = row[i];
      if (!header) continue;

      const sectionName = getSectionName(i);

      if (isImageHeader(header)) {
        // This is an image column
        if (isImageUrl(value)) {
          // Find preceding Q column (skip other image cols going left)
          let precQ = '';
          let precA = '';
          for (let j = i - 1; j >= COL.QUESTIONS_START; j--) {
            if (!isImageHeader(headers[j])) {
              precQ = headers[j] || '';
              precA = row[j] != null ? String(row[j]).trim() : '';
              break;
            }
          }
          imageEntries.push({ sectionName, imageUrl: value, imageHeader: header, precedingQuestion: precQ, precedingAnswer: precA });

          // Also attach imageUrl to the preceding QAPair if already added
          const pairs = sectionMap.get(sectionName) || [];
          const lastPair = pairs[pairs.length - 1];
          if (lastPair && !lastPair.imageUrl) {
            lastPair.imageUrl = value;
          }
        }
      } else {
        // Regular question column
        if (value != null && String(value).trim() !== '') {
          const pairs = sectionMap.get(sectionName) || [];
          pairs.push({ question: header, answer: String(value).trim() });
          sectionMap.set(sectionName, pairs);
        }
      }
    }

    const sections: SurveySection[] = Array.from(sectionMap.entries()).map(([name, qaPairs]) => ({ name, qaPairs }));

    return {
      id: get(0),
      email: get(COL.EMAIL),
      firstName: get(COL.FIRST_NAME),
      lastName: get(COL.LAST_NAME),
      fullName: `${get(COL.FIRST_NAME)} ${get(COL.LAST_NAME)}`.trim(),
      store: get(COL.STORE),
      storeCode: get(COL.STORE_CODE),
      province: get(COL.PROVINCE),
      date: dateStr,
      dateRaw: dateObj,
      dayOfWeek: get(COL.DAY),
      sections,
      imageEntries,
    };
  });

  // Aggregates
  const uniqueEmails = [...new Set(rows.map(r => r.email).filter(Boolean))];
  const uniqueStores = [...new Set(rows.map(r => r.store).filter(Boolean))];
  const uniqueDays = [...new Set(rows.map(r => r.dayOfWeek).filter(Boolean))];

  const dates = rows.map(r => r.dateRaw).filter((d): d is Date => d != null);
  dates.sort((a, b) => a.getTime() - b.getTime());
  const dateRange = {
    from: dates.length ? formatDateObj(dates[0]) : '',
    to: dates.length ? formatDateObj(dates[dates.length - 1]) : '',
  };

  return { rows, uniqueUsers: uniqueEmails, uniqueStores, uniqueDays, dateRange, totalRows: rows.length };
}

export function buildUserSummaries(data: ParsedData): UserSummary[] {
  const map = new Map<string, UserSummary>();
  for (const row of data.rows) {
    if (!row.email) continue;
    const existing = map.get(row.email);
    if (!existing) {
      map.set(row.email, {
        fullName: row.fullName,
        email: row.email,
        uniqueStores: 0,
        totalSurveys: 0,
        dayCounts: {},
      });
    }
    const s = map.get(row.email)!;
    s.totalSurveys += 1;
    if (row.dayOfWeek) {
      s.dayCounts[row.dayOfWeek] = (s.dayCounts[row.dayOfWeek] || 0) + 1;
    }
  }

  // Unique stores per user
  for (const email of map.keys()) {
    const userRows = data.rows.filter(r => r.email === email);
    const stores = new Set(userRows.map(r => r.store).filter(Boolean));
    map.get(email)!.uniqueStores = stores.size;
  }

  return Array.from(map.values());
}
