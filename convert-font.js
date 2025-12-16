#!/usr/bin/env node

/**
 * Font Converter Script
 * Converts TTF/OTF font files to Base64 format for use in fonts.ts
 * 
 * Usage:
 *   node convert-font.js path/to/font.ttf
 *   node convert-font.js path/to/font-bold.ttf
 * 
 * Output:
 *   Prints Base64 string to console and saves to font-base64.txt
 */

const fs = require('fs');
const path = require('path');

// Get font file path from command line argument
const fontPath = process.argv[2];

if (!fontPath) {
  console.error('❌ Error: Please provide a font file path');
  console.error('\nUsage:');
  console.error('  node convert-font.js path/to/font.ttf');
  console.error('  node convert-font.js path/to/font-bold.ttf');
  process.exit(1);
}

// Check if file exists
if (!fs.existsSync(fontPath)) {
  console.error(`❌ Error: File not found: ${fontPath}`);
  process.exit(1);
}

try {
  // Read font file
  const fontData = fs.readFileSync(fontPath);
  
  // Convert to Base64
  const base64String = fontData.toString('base64');
  
  // Get file name for reference
  const fileName = path.basename(fontPath);
  const fileSize = (fontData.length / 1024).toFixed(2); // KB
  const base64Size = (base64String.length / 1024).toFixed(2); // KB
  
  // Print to console
  console.log('\n' + '='.repeat(80));
  console.log('✅ Font Conversion Successful');
  console.log('='.repeat(80));
  console.log(`\n📄 Font File: ${fileName}`);
  console.log(`📦 Original Size: ${fileSize} KB`);
  console.log(`🔤 Base64 Size: ${base64Size} KB`);
  console.log(`\n📝 Base64 String (copy this to fonts.ts):\n`);
  console.log(base64String);
  console.log('\n' + '='.repeat(80) + '\n');
  
  // Save to file
  const outputPath = path.join(process.cwd(), 'font-base64.txt');
  fs.writeFileSync(outputPath, base64String);
  console.log(`✅ Saved to: ${outputPath}\n`);
  
  // Instructions
  console.log('📌 Next Steps:');
  console.log('1. Open fonts.ts');
  console.log('2. Find the appropriate placeholder (primaryRegular or primaryBold)');
  console.log('3. Replace the placeholder with the Base64 string above');
  console.log('4. Save fonts.ts');
  console.log('5. Build and test: npm run build\n');
  
  // Additional info
  console.log('💡 Tips:');
  console.log('- For regular font → customFonts.primaryRegular.data');
  console.log('- For bold font → customFonts.primaryBold.data');
  console.log('- Keep the font name consistent (e.g., "CustomFont-Regular")');
  console.log('- Test the PDF output to verify font is applied\n');
  
} catch (error) {
  console.error('❌ Error converting font:', error.message);
  process.exit(1);
}
