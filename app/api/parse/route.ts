import { NextRequest, NextResponse } from 'next/server';
import { parseExcel, buildUserSummaries } from '@/lib/excel-parser';

export const maxDuration = 30;

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });

    const buffer = await file.arrayBuffer();
    const data = parseExcel(buffer);
    const userSummaries = buildUserSummaries(data);

    const totalImages = data.rows.reduce((sum, r) => sum + r.imageEntries.length, 0);

    return NextResponse.json({
      totalRows: data.totalRows,
      uniqueUsers: data.uniqueUsers,
      uniqueStores: data.uniqueStores,
      uniqueDays: data.uniqueDays,
      dateRange: data.dateRange,
      totalImages,
      rows: data.rows,
      userSummaries,
    });
  } catch (err) {
    console.error('Parse error:', err);
    return NextResponse.json({ error: 'Failed to parse Excel file' }, { status: 500 });
  }
}
