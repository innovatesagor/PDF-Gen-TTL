import fs from 'fs';
import path from 'path';

const projectDir = '/Users/sagor/Downloads/tusuka-inventory-report-generator (5)';
const regularPath = path.join(projectDir, 'public/fonts/CustomFont-Regular.ttf');
const boldPath = path.join(projectDir, 'public/fonts/CustomFont-Bold.ttf');

console.log('Reading font files...');
const regularBuffer = fs.readFileSync(regularPath);
const boldBuffer = fs.readFileSync(boldPath);

const regularBase64 = regularBuffer.toString('base64');
const boldBase64 = boldBuffer.toString('base64');

const content = `/**
 * Embedded Custom Fonts (Base64)
 * Auto-generated from TTF files
 * Run: node generate-fonts.mjs
 */

const FONT_REGULAR_BASE64 = '${regularBase64}';
const FONT_BOLD_BASE64 = '${boldBase64}';

export const registerPDFFonts = async (jsPDFInstance: any): Promise<void> => {
  try {
    // Register fonts with jsPDF using embedded base64
    jsPDFInstance.addFileToVFS('CustomFont-Regular.ttf', FONT_REGULAR_BASE64);
    jsPDFInstance.addFileToVFS('CustomFont-Bold.ttf', FONT_BOLD_BASE64);
    
    jsPDFInstance.addFont('CustomFont-Regular.ttf', 'CustomFont', 'normal');
    jsPDFInstance.addFont('CustomFont-Bold.ttf', 'CustomFont', 'bold');
    
    console.log('✅ Custom fonts embedded in PDF');
  } catch (error) {
    console.error('❌ Font registration error:', error);
  }
};

export const getActiveFontName = (): string => {
  return 'CustomFont';
};

export const getFallbackFontName = (): string => {
  return 'helvetica';
};
`;

fs.writeFileSync(path.join(projectDir, 'fonts.ts'), content);
console.log('✅ fonts.ts updated with embedded fonts');
