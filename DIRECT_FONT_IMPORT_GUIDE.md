# 🎨 Direct Font File Import - Quick Setup Guide

## ✅ YES! You Can Use TTF/OTF/WOFF Files Directly

**No Base64 conversion needed!** Just copy your font files and reference them.

---

## 📂 Setup in 3 Steps

### Step 1: Create Font Folder
```bash
mkdir -p public/fonts
```

### Step 2: Copy Your Font Files
Copy your font files to `public/fonts/`:
```
public/
└── fonts/
    ├── CustomFont-Regular.ttf    ← Your regular font
    ├── CustomFont-Bold.ttf       ← Your bold font
    ├── CustomFont-Regular.otf    ← Or use OTF
    └── CustomFont-Bold.otf       ← Or use OTF
```

**Supported Formats:**
- ✅ TTF (TrueType) - Most compatible
- ✅ OTF (OpenType) - Also good
- ✅ WOFF (Web Open Font Format) - Modern web
- ✅ WOFF2 (Compressed WOFF) - Best for web

### Step 3: Update Font Paths in `fonts.ts`
Edit the `fontPaths` object:

```typescript
export const fontPaths = {
  primaryRegular: '/fonts/CustomFont-Regular.ttf',  // Update filename
  primaryBold: '/fonts/CustomFont-Bold.ttf',        // Update filename
  fallback: 'helvetica',
};
```

**That's it! Done!** ✅

---

## 🎯 Usage in Your Project

### For Web UI (HTML/CSS)

**Option 1: Auto-inject CSS**

In `index.tsx`:
```typescript
import { getFontFaceCSS } from './fonts';

// Inject custom font CSS
const fontCSS = getFontFaceCSS();
const styleSheet = document.createElement('style');
styleSheet.textContent = fontCSS;
document.head.appendChild(styleSheet);
```

**Option 2: Manual CSS**

In `index.css`:
```css
@font-face {
  font-family: 'CustomFont';
  src: url('/fonts/CustomFont-Regular.ttf') format('truetype');
  font-weight: normal;
}

@font-face {
  font-family: 'CustomFont';
  src: url('/fonts/CustomFont-Bold.ttf') format('truetype');
  font-weight: bold;
}

body {
  font-family: 'CustomFont', sans-serif;
}
```

### For PDF Generation

In `services/reportService.ts`:
```typescript
import { registerPDFFonts, getActiveFontName } from '../fonts';

const generatePDF = async (header: ReportHeader, items: LineItem[], totals: Totals, filename: string) => {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  
  // Register custom fonts
  await registerPDFFonts(doc);
  
  // Use custom font
  const fontName = getActiveFontName();
  doc.setFont(fontName, 'normal');
  doc.setFontSize(18);
  doc.text("Tusuka Trousers Ltd.", pageWidth / 2, headerY, { align: 'center' });
  
  // ... rest of PDF generation
};
```

---

## 📊 Comparison: Methods

| Method | Easy | No Base64 | Web | PDF | Speed |
|--------|------|-----------|-----|-----|-------|
| **Direct Files** | ⭐⭐⭐⭐⭐ | ✅ Yes | ✅ | ✅ | ⭐⭐⭐⭐⭐ |
| **Base64** | ⭐⭐ | ❌ No | ✅ | ✅ | ⭐⭐⭐ |
| **CDN (Google)** | ⭐⭐⭐ | ✅ Yes | ✅ | ⭐⭐ | ⭐⭐⭐ |

---

## 🔧 File Formats Guide

### TTF (TrueType Font)
```typescript
// Use this CSS format:
@font-face {
  src: url('/fonts/font.ttf') format('truetype');
}
```
- ✅ Works in all browsers
- ✅ Works in PDFs
- ✅ Good file size
- **Best for:** Desktop and PDF

### OTF (OpenType Font)
```typescript
// Use this CSS format:
@font-face {
  src: url('/fonts/font.otf') format('opentype');
}
```
- ✅ Modern format
- ✅ Better features than TTF
- ✅ Works in most browsers
- **Best for:** Professional designs

### WOFF / WOFF2 (Web Fonts)
```typescript
// Use this CSS format:
@font-face {
  src: url('/fonts/font.woff2') format('woff2'),
       url('/fonts/font.woff') format('woff');
}
```
- ✅ Compressed, smaller files
- ✅ Optimized for web
- ❌ May not work in all PDFs
- **Best for:** Web UI only

---

## 💡 Pro Tips

### 1. Use TTF for Both Web and PDF
TTF works everywhere and is the safest choice:
```typescript
fontPaths = {
  primaryRegular: '/fonts/MyFont-Regular.ttf',
  primaryBold: '/fonts/MyFont-Bold.ttf',
};
```

### 2. Reduce Font File Size
Use online tools to subset fonts (keep only needed characters):
- https://transfonter.org/ (add `subsetting`)
- Reduces file from 500KB to 50KB!

### 3. Use WOFF2 for Modern Browsers
For web-only fonts, WOFF2 is 30% smaller:
```typescript
// Update in fonts.ts to use WOFF2
primaryRegular: '/fonts/MyFont-Regular.woff2',
```

### 4. Fallback Font
Always keep the fallback for when custom font fails:
```typescript
fallback: 'helvetica',  // Falls back to helvetica if font fails
```

---

## ✅ Complete Example

### Project Structure
```
my-project/
├── public/
│   └── fonts/
│       ├── Roboto-Regular.ttf
│       └── Roboto-Bold.ttf
├── fonts.ts                  ← Font config
├── index.tsx                 ← Inject CSS
├── index.css                 ← Use font
├── services/
│   └── reportService.ts      ← Use in PDF
└── App.tsx
```

### fonts.ts
```typescript
export const fontPaths = {
  primaryRegular: '/fonts/Roboto-Regular.ttf',
  primaryBold: '/fonts/Roboto-Bold.ttf',
  fallback: 'helvetica',
};

export const getFontFaceCSS = (): string => {
  return `
    @font-face {
      font-family: 'CustomFont';
      src: url('${fontPaths.primaryRegular}') format('truetype');
      font-weight: normal;
    }
    @font-face {
      font-family: 'CustomFont';
      src: url('${fontPaths.primaryBold}') format('truetype');
      font-weight: bold;
    }
  `;
};
```

### index.tsx
```typescript
import { getFontFaceCSS } from './fonts';

const fontCSS = getFontFaceCSS();
const styleSheet = document.createElement('style');
styleSheet.textContent = fontCSS;
document.head.appendChild(styleSheet);
```

### index.css
```css
body {
  font-family: 'CustomFont', sans-serif;
  font-size: 14px;
}

h1, h2, h3 {
  font-weight: bold;
}
```

### reportService.ts
```typescript
import { registerPDFFonts, getActiveFontName } from '../fonts';

const generatePDF = async (...) => {
  const doc = new jsPDF({ orientation: 'landscape' });
  await registerPDFFonts(doc);
  
  doc.setFont(getActiveFontName(), 'normal');
  // ... rest of PDF code
};
```

---

## 🚀 Deploy & Test

```bash
# Build
npm run build

# Test locally
npm run dev
# Visit http://localhost:3000

# Check font loaded
# Open DevTools → Network tab → look for fonts folder requests
```

---

## 🎉 Done!

Your fonts are now:
- ✅ Used in Web UI
- ✅ Used in PDF generation
- ✅ No Base64 conversion needed
- ✅ Fast and efficient
- ✅ Production ready!

**Enjoy your custom fonts!** 🚀
