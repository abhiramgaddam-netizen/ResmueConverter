import { NextRequest, NextResponse } from 'next/server';
import { resolveAvailableModel } from '@/lib/gemini';

export async function POST(req: NextRequest) {
  try {
    const { apiKey } = await req.json();
    const activeKey = apiKey || process.env.GEMINI_API_KEY;

    if (!activeKey) {
      return NextResponse.json(
        { valid: false, error: 'No authorization key provided.' },
        { status: 400 }
      );
    }

    const { name, apiVersion } = await resolveAvailableModel(activeKey);

    return NextResponse.json({
      valid: true,
      model: name,
      apiVersion,
      message: 'Transformation Engine successfully connected and verified.',
    });
  } catch (error: any) {
    return NextResponse.json(
      {
        valid: false,
        error: error?.message || 'Could not validate engine authorization key.',
      },
      { status: 400 }
    );
  }
}
