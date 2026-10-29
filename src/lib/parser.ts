import mammoth from 'mammoth';
import { extractText } from 'unpdf';

export async function extractTextFromFile(
  buffer: Buffer,
  filename: string
): Promise<string> {
  const lowerName = filename.toLowerCase();

  if (lowerName.endsWith('.pdf')) {
    const uint8 = new Uint8Array(buffer);
    const { text } = await extractText(uint8);
    return Array.isArray(text) ? text.join('\n\n') : String(text || '');
  }

  if (lowerName.endsWith('.docx') || lowerName.endsWith('.doc')) {
    const result = await mammoth.extractRawText({ buffer });
    return result.value || '';
  }

  if (lowerName.endsWith('.txt') || lowerName.endsWith('.md')) {
    return buffer.toString('utf-8');
  }

  throw new Error(`Unsupported file format: ${filename}. Please upload .pdf, .docx, or .txt.`);
}
