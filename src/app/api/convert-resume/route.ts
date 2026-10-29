import { NextRequest, NextResponse } from 'next/server';
import { convertCandidateResume, extractNameFromFilename } from '@/lib/gemini';
import { generateDocxBuffer, detectTemplateHeaderType } from '@/lib/docx-generator';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      candidateText,
      targetFormatText,
      filename,
      apiKey,
      templateDocxBase64,
      removePersonalInfo = true,
    } = body;

    if (!candidateText || !candidateText.trim()) {
      return NextResponse.json(
        { error: 'Candidate resume text is required.' },
        { status: 400 }
      );
    }

    if (!targetFormatText || !targetFormatText.trim()) {
      return NextResponse.json(
        { error: 'Target resume format specification is required.' },
        { status: 400 }
      );
    }

    // Convert candidate resume into structured facts (with optional personal detail removal)
    const structuredData = await convertCandidateResume({
      candidateText,
      targetFormatText,
      filename,
      apiKey,
      removePersonalInfo: Boolean(removePersonalInfo),
    });

    if (removePersonalInfo) {
      structuredData.contactInfo = {};
    }

    // Ensure candidate name is always non-empty
    if (!structuredData.fullName || structuredData.fullName.trim() === '' || structuredData.fullName === 'Candidate Profile') {
      structuredData.fullName = filename ? extractNameFromFilename(filename) : 'Candidate';
    }

    // If a custom template .docx was uploaded, use its buffer, else use the default Tecnics template
    const templateBuffer = templateDocxBase64
      ? Buffer.from(templateDocxBase64, 'base64')
      : null;

    // Generate Word Document using the actual template preserving logo and layout
    const docxBuffer = await generateDocxBuffer(structuredData, templateBuffer, {
      removePersonalInfo: Boolean(removePersonalInfo),
    });
    const docxBase64 = docxBuffer.toString('base64');
    const headerType = detectTemplateHeaderType(templateBuffer);

    return NextResponse.json({
      success: true,
      candidateName: structuredData.fullName,
      targetRole: structuredData.targetRole,
      headerType,
      docxBase64,
      data: structuredData,
    });
  } catch (error: any) {
    console.error('Error in /api/convert-resume:', error);
    return NextResponse.json(
      { error: error?.message || 'Failed to standardize resume.' },
      { status: 500 }
    );
  }
}
