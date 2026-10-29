import { NextRequest, NextResponse } from 'next/server';
import { extractTextFromFile } from '@/lib/parser';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'No file uploaded.' }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const text = await extractTextFromFile(buffer, file.name);

    if (!text || text.trim().length === 0) {
      return NextResponse.json(
        { error: `Could not extract text from "${file.name}". Ensure the file is not an empty or password-protected document.` },
        { status: 422 }
      );
    }

    return NextResponse.json({
      success: true,
      filename: file.name,
      textLength: text.length,
      text,
    });
  } catch (error: any) {
    console.error('Error in /api/extract-text:', error);
    return NextResponse.json(
      { error: error?.message || 'Failed to parse document text.' },
      { status: 500 }
    );
  }
}
