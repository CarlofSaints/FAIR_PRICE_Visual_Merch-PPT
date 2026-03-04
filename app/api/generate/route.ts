import { NextRequest, NextResponse } from 'next/server';
import { parseExcel, buildUserSummaries } from '@/lib/excel-parser';
import { buildPptx } from '@/lib/ppt-builder';

export const maxDuration = 60;

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });

    // Read the Perigee session cookie stored after login
    const perigeeCookie = req.cookies.get('perigee_session')?.value ?? null;
    if (!perigeeCookie) {
      return NextResponse.json({ error: 'Not logged in to Perigee. Please log in first.' }, { status: 401 });
    }

    const buffer = await file.arrayBuffer();
    const data = parseExcel(buffer);
    const summaries = buildUserSummaries(data);

    const pptxBuffer = await buildPptx(data, summaries, perigeeCookie);

    return new NextResponse(new Uint8Array(pptxBuffer), {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="FairPrice_VisualMerch.pptx"',
        'Content-Length': pptxBuffer.length.toString(),
      },
    });
  } catch (err) {
    console.error('Generate error:', err);
    return NextResponse.json({ error: 'Failed to generate presentation' }, { status: 500 });
  }
}
