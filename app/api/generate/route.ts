import { NextRequest, NextResponse } from 'next/server';
import { parseExcel, buildUserSummaries } from '@/lib/excel-parser';
import { buildPptx } from '@/lib/ppt-builder';

export const maxDuration = 60;

export async function POST(req: NextRequest) {
  // Only available when running locally — server has SA IP so Perigee isn't blocked.
  // On Vercel all datacenter IPs are blocked; use client-side generation instead.
  const host = req.headers.get('host') || '';
  const isLocal = host.startsWith('localhost') || host.startsWith('127.0.0.1');
  if (!isLocal) {
    return NextResponse.json({ error: 'Server-side generation is only available locally.' }, { status: 403 });
  }

  try {
    const form = await req.formData();
    const file = form.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });

    const buffer = await file.arrayBuffer();
    const data = parseExcel(buffer);
    const summaries = buildUserSummaries(data);

    // Images are public — no Perigee session cookie needed from localhost
    const pptxBuffer = await buildPptx(data, summaries, '');

    return new NextResponse(new Uint8Array(pptxBuffer), {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="FairPrice_VisualMerch.pptx"',
      },
    });
  } catch (err) {
    console.error('Generate error:', err);
    return NextResponse.json({ error: 'Failed to generate presentation' }, { status: 500 });
  }
}
