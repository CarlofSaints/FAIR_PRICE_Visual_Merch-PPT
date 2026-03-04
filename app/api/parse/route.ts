import { NextRequest, NextResponse } from 'next/server';
import { parseExcel } from '@/lib/excel-parser';

export const maxDuration = 30;

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });

    const buffer = await file.arrayBuffer();
    const data = parseExcel(buffer);

    return NextResponse.json({
      totalRows: data.totalRows,
      uniqueUsers: data.uniqueUsers.length,
      uniqueStores: data.uniqueStores.length,
      uniqueDays: data.uniqueDays,
      dateRange: data.dateRange,
    });
  } catch (err) {
    console.error('Parse error:', err);
    return NextResponse.json({ error: 'Failed to parse Excel file' }, { status: 500 });
  }
}
