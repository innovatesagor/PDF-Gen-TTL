# 🎨 Custom Fonts Guide

## Overview

This guide explains how to add your custom fonts to both the **Web UI** and **PDF** generation.

---

## 📝 Step-by-Step Instructions

### Step 1: Prepare Your Font Files

You need font files in these formats:
- **TTF** (TrueType Font) - Recommended for both web and PDF
- **OTF** (OpenType Font) - Also supported
- **WOFF2** - For web only (convert from TTF/OTF)

### Step 2: Convert Font to Base64

Use Node.js to convert your font file to Base64:

```bash
# Create a conversion script
cat > convert-font.js << 'EOF'
const fs = require('fs');
const path = require('path');

// Replace with your font file path
const fontPath = path.join(__dirname, 'path/to/your-font.ttf');
const fontData = fs.readFileSync(fontPath);
const base64 = fontData.toString('base64');

console.log('Base64 Font String:');
console.log(base64);

// Also save to file for reference
fs.writeFileSync('font-base64.txt', base64);
console.log('\n✅ Saved to font-base64.txt');
EOF

node convert-font.js
```

### Step 3: Add Font to `fonts.ts`

Open `fonts.ts` and replace the placeholder with your Base64 string:

```typescript
export const customFonts = {
  // Regular Font
  primaryRegular: {
    name: 'CustomFont-Regular',
    data: 'PASTE_YOUR_BASE64_STRING_HERE', // ← Replace this
    style: 'normal',
  },

  // Bold Font
  primaryBold: {
    name: 'CustomFont-Bold',
    data: 'PASTE_YOUR_BASE64_BOLD_STRING_HERE', // ← Replace this
    style: 'bold',
  },
  // ... rest of config
};
```

### Step 4: Register Fonts in PDF Generation

Update `services/reportService.ts` to use custom fonts:

```typescript
import { registerPDFFonts, getActiveFontName } from '../fonts';

const generatePDF = (header: ReportHeader, items: LineItem[], totals: Totals, filename: string) => {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  
  // Register custom fonts
  registerPDFFonts(doc);
  
  // Use custom font in PDF
  const fontName = getActiveFontName();
  doc.setFont(fontName, 'normal');
  doc.setFontSize(18);
  doc.text("Tusuka Trousers Ltd.", pageWidth / 2, headerY, { align: 'center' });
  
  // ... rest of PDF generation
};
```

### Step 5: Use Font in Web UI

Update `index.tsx` to inject font CSS:

```typescript
import { getFontFaceCSS } from './fonts';
import './index.css';

// Inject custom font CSS
const fontCSS = getFontFaceCSS();
const styleSheet = document.createElement('style');
styleSheet.textContent = fontCSS;
document.head.appendChild(styleSheet);

// Now update App.tsx or index.css to use the font
```

Update `index.css`:

```css
body {
  font-family: 'CustomFont', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  /* ... other styles */
}
```

---

## 📂 File Structure After Setup

```
project/
├── fonts.ts                          ← Your custom fonts (Base64)
├── services/
│   └── reportService.ts              ← Uses fonts via imports
├── index.tsx                         ← Injects font CSS
├── index.css                         ← Uses custom font
└── convert-font.js                   ← Helper script (optional)
```

---

## 🔄 Workflow for Adding Custom Fonts

### For Regular Font (e.g., "Roboto-Regular.ttf")

1. Convert to Base64 → copy output
2. Paste in `fonts.ts` → `customFonts.primaryRegular.data`
3. Font automatically available in PDF and Web UI

### For Bold Font (e.g., "Roboto-Bold.ttf")

1. Convert to Base64 → copy output
2. Paste in `fonts.ts` → `customFonts.primaryBold.data`
3. Font automatically available in PDF and Web UI

---

## 💡 Example: Using "Inter" Font

If you want to use the "Inter" font:

1. Download `Inter-Regular.ttf` and `Inter-Bold.ttf`
2. Convert both to Base64
3. Add to `fonts.ts`:

```typescript
export const customFonts = {
  primaryRegular: {
    name: 'Inter-Regular',
    data: 'AAEAAAALAIAAAwAwT1MvMg8SBfMAAAC8AAAAYGNtYXAXV/...',
    style: 'normal',
  },
  primaryBold: {
    name: 'Inter-Bold',
    data: 'AAEAAAALAIAAAwAwT1MvMg8SBfMAAAC8AAAAYGNtYXAXV/...',
    style: 'bold',
  },
  // ...
};
```

4. Update CSS and PDF generation to use "Inter-Regular"

---

## ⚙️ Advanced: Dynamic Font Loading

If you want to load fonts dynamically:

```typescript
export const loadFontDynamically = async (fontName: string): Promise<void> => {
  const response = await fetch(`/fonts/${fontName}.ttf`);
  const fontData = await response.arrayBuffer();
  const base64 = btoa(String.fromCharCode(...new Uint8Array(fontData)));
  
  customFonts.primaryRegular.data = base64;
  // Re-inject CSS or register with PDF
};
```

---

## 🎯 Current Status

**Files Created:**
- ✅ `fonts.ts` - Font management system

**Next Steps:**
1. Prepare your custom font files (TTF/OTF)
2. Convert to Base64 using the script provided
3. Add Base64 strings to `fonts.ts`
4. Update `reportService.ts` to use custom fonts
5. Update `index.css` to use custom fonts

---

## 🔗 Useful Resources

- **Font Conversion Tool**: https://transfonter.org/ (online, no installation needed)
- **Base64 Encoder**: https://www.base64encode.org/
- **Font Formats**: TTF (best support), OTF (also good), WOFF2 (web only)

---

## ✅ Checklist

- [ ] Custom font files prepared (TTF/OTF)
- [ ] Base64 strings generated
- [ ] `fonts.ts` updated with font data
- [ ] `reportService.ts` updated to import and use fonts
- [ ] `index.tsx` updated to inject CSS
- [ ] `index.css` updated with font-family
- [ ] Build and test ✅

---

**Ready to add your custom fonts!** 🚀
