import * as fs from 'fs';
import * as path from 'path';

/**
 * Generate a file number page PDF for separator printing
 * @param fileNumber - File number (1-10)
 */
export async function generateFileNumberPage(fileNumber: number): Promise<Buffer> {
  // Create a PDF with "File no: X" centered on the page
  const fileNumberText = `File no: ${fileNumber}`;
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Resources <<
/Font <<
/F1 <<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica-Bold
>>
>>
>>
/Contents 4 0 R
>>
endobj
4 0 obj
<<
/Length 120
>>
stream
BT
/F1 24 Tf
200 400 Td
(${fileNumberText}) Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000306 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
446
%%EOF`;

  return Buffer.from(pdfContent);
}

/**
 * Generate a letter separator page PDF
 * @param letter - Letter to print (A-Z)
 */
export async function generateLetterSeparator(letter: string): Promise<Buffer> {
  // Create a PDF with a large letter centered on the page
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Resources <<
/Font <<
/F1 <<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica-Bold
>>
>>
>>
/Contents 4 0 R
>>
endobj
4 0 obj
<<
/Length 100
>>
stream
BT
/F1 120 Tf
200 400 Td
(${letter}) Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000306 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
446
%%EOF`;

  return Buffer.from(pdfContent);
}

/**
 * Save file to temporary directory
 */
export function saveTempFile(data: Buffer, filename: string): string {
  const tempDir = path.join(process.cwd(), 'temp');
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }
  const filePath = path.join(tempDir, filename);
  fs.writeFileSync(filePath, data);
  return filePath;
}

/**
 * Clean up temporary file
 */
export function cleanupTempFile(filePath: string): void {
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (error) {
    console.error(`Error cleaning up temp file ${filePath}:`, error);
  }
}

