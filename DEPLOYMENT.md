# 🚀 Deployment Guide - Tusuka Inventory Report Generator

Complete step-by-step guide to deploy your app to Vercel.

---

## 📋 Pre-Deployment Checklist

- [ ] Node.js 16+ installed locally
- [ ] npm or yarn installed
- [ ] Vercel account created (free at vercel.com)
- [ ] GitHub account (for Git deployment)
- [ ] Git installed and configured

---

## Method 1: Vercel CLI (Fastest & Easiest)

### Step 1: Install Vercel CLI

```bash
npm install -g vercel
```

### Step 2: Build Locally (Optional but Recommended)

```bash
npm install
npm run build
```

### Step 3: Deploy

```bash
vercel
```

### Step 4: Follow Prompts

- **Project name**: tusuka-inventory-report-generator
- **Project root**: ./ (current directory)
- **Build command**: npm run build (default)
- **Output directory**: dist (default)

```
✓ Linked to [your-username]/tusuka-inventory-report-generator
✓ Production: https://tusuka-inventory-report-generator.vercel.app
✓ Deploy complete!
```

**Done! 🎉 Your app is live in ~1 minute**

---

## Method 2: Git Push Deployment (GitHub)

### Step 1: Initialize Git Repository

```bash
git init
git add .
git commit -m "Initial commit - Vercel deployment ready"
```

### Step 2: Push to GitHub

```bash
git remote add origin https://github.com/YOUR_USERNAME/tusuka-inventory-report-generator.git
git branch -M main
git push -u origin main
```

### Step 3: Connect to Vercel Dashboard

1. Go to [vercel.com](https://vercel.com)
2. Click **"New Project"**
3. Select **"Import Git Repository"**
4. Paste GitHub URL and click **Import**
5. Click **Deploy**

```
✓ Repository connected
✓ Build successful
✓ Deployment complete!
```

**Your app is live in ~2-3 minutes**

---

## Method 3: Manual Folder Upload

### Step 1: Build the Project

```bash
npm install
npm run build
```

### Step 2: Upload to Vercel

1. Go to [vercel.com](https://vercel.com)
2. Click **"New Project"**
3. Select **"Other"** → **"Create a Git Repository"**
4. Drag & drop the entire project folder
5. Click **Deploy**

**Your app is live in ~2 minutes**

---

## 🔧 Configure Environment Variables (Optional)

### For Vercel Dashboard:

1. Go to your **Project Settings**
2. Select **"Environment Variables"**
3. Add `GEMINI_API_KEY=your_key_here`
4. Click **"Save"**
5. **Redeploy** the project

### Redeploy After Changes:

```bash
vercel --prod
```

---

## ✅ Post-Deployment Verification

### Check Deployment Status

1. Go to [vercel.com/dashboard](https://vercel.com/dashboard)
2. Click on **tusuka-inventory-report-generator**
3. Check **Deployments** tab
4. Should see **"Ready"** status ✓

### Test Your Live App

1. Click the **Production URL**
2. Fill in a test bill with sample data
3. Click **"Generate Report (PDF & Excel)"**
4. Verify PDF and Excel files download

### Verify Performance

Use Vercel Analytics:
1. Project Settings → Analytics
2. Check Real User Monitoring (RUM)
3. Review performance metrics

---

## 🔄 Auto-Deployment on Git Push

Once connected to Vercel:

1. Make local changes
2. Commit and push to GitHub:
   ```bash
   git add .
   git commit -m "Add feature XYZ"
   git push origin main
   ```
3. Vercel **automatically deploys** within 1-2 minutes!

---

## 🐛 Troubleshooting

### "Build Failed" Error

**Check:** 
- Node.js version compatibility
- All dependencies in package.json
- No TypeScript errors

**Solution:**
```bash
npm install
npm run build
```

### "Deployment Stuck"

**Check:**
- Build logs in Vercel dashboard
- Large files being uploaded
- Network issues

**Solution:**
- Click **"Redeploy"** button
- Try deploying again with Vercel CLI

### "Not Found" (404) Error

**Issue:** SPA routing issue

**Solution:** Vercel.json should have:
```json
{
  "buildCommand": "npm run build",
  "framework": "vite",
  "outputDirectory": "dist"
}
```

### "Environment Variables Not Working"

**Solution:**
1. Verify variables in Vercel dashboard
2. Redeploy the project
3. Clear browser cache (Ctrl+Shift+Del)

---

## 📊 Deployment Monitoring

### Real-Time Logs

```bash
vercel logs
```

### View All Deployments

```bash
vercel list
```

### Rollback to Previous Deployment

```bash
vercel rollback
```

---

## 🔐 Security Best Practices

- ✅ Never commit .env.local to Git
- ✅ Use Vercel dashboard for secrets
- ✅ Enable HTTPS (automatic on Vercel)
- ✅ Keep dependencies updated
- ✅ Monitor Vercel analytics for suspicious activity

---

## 📈 Monitoring & Analytics

### Enable Analytics in Vercel

1. Project Settings → Analytics
2. Enable **Real User Monitoring (RUM)**
3. Monitor:
   - Page load times
   - User interactions
   - Error rates
   - Traffic patterns

### Performance Tips

- [ ] Use Vercel's Image Optimization
- [ ] Enable Gzip compression (automatic)
- [ ] Monitor Core Web Vitals
- [ ] Review build time (target: <30s)

---

## 🚀 Custom Domain (Optional)

### Add Custom Domain

1. Project Settings → Domains
2. Click **"Add"**
3. Enter your domain (e.g., reports.tusuka.com)
4. Follow DNS configuration steps
5. Wait for DNS propagation (5-30 minutes)

---

## 📞 Support

- **Vercel Docs**: https://vercel.com/docs
- **Vite Docs**: https://vitejs.dev
- **React Docs**: https://react.dev
- **Community Help**: Vercel Discord

---

## 🎯 Next Steps After Deployment

1. ✅ Test all features on live URL
2. ✅ Share with team members
3. ✅ Monitor performance metrics
4. ✅ Gather feedback
5. ✅ Plan updates and improvements

---

**Your app is now production-ready and deployed! 🎉**

For questions or issues, refer to the main README.md or Vercel documentation.

*Last Updated: December 2025*
