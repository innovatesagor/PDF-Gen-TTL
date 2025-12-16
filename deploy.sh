#!/bin/bash
# Tusuka Inventory Report Generator - Quick Deployment Script

echo "🚀 Tusuka Inventory Report Generator - Vercel Deployment"
echo "==========================================================="
echo ""

# Check Node.js
echo "✅ Checking Node.js..."
node --version

# Install dependencies
echo "✅ Installing dependencies..."
npm install

# Build
echo "✅ Building application..."
npm run build

# Check if build succeeded
if [ $? -eq 0 ]; then
  echo ""
  echo "✅ Build successful!"
  echo ""
  echo "========== DEPLOYMENT OPTIONS =========="
  echo ""
  echo "1️⃣  Deploy to Vercel using CLI (Recommended):"
  echo "   npm install -g vercel"
  echo "   vercel"
  echo ""
  echo "2️⃣  Deploy to Vercel using Git:"
  echo "   git push origin main"
  echo "   (Go to vercel.com/new and select your repository)"
  echo ""
  echo "3️⃣  Test locally before deploying:"
  echo "   npm run preview"
  echo ""
  echo "=========================================="
  echo ""
  echo "📦 Build Output:"
  echo "   - Location: dist/"
  echo "   - Size: $(du -sh dist | cut -f1)"
  echo ""
else
  echo "❌ Build failed! Please check errors above."
  exit 1
fi
