# 🚀 Render.com Deployment Checklist

## ✅ Already Done (by you or me)

- ✅ `render.yaml` - Deployment configuration
- ✅ `build.sh` - Build script for dependencies
- ✅ `start-render.sh` - Start script with Google credentials handling
- ✅ `requirements.txt` - All Python dependencies including gunicorn
- ✅ `.gitignore` - Credentials and output folders excluded
- ✅ `frontend/index.html` - Uses relative API URLs (works on Render)
- ✅ Git repository ready to push

---

## 📝 What YOU Need to Do

### Step 1: Push to GitHub ⏱️ 2 minutes

```powershell
cd "c:\Users\User\Desktop\Reports Automation\Python Engine"
git status  # Verify everything is clean
git push origin main  # Push to GitHub
```

**⚠️ IMPORTANT:** Make sure your `Credentials/` folder is NOT pushed to GitHub (it's already in `.gitignore`)

---

### Step 2: Sign Up on Render.com ⏱️ 2 minutes

1. Go to: **https://render.com**
2. Click **"Get Started"**
3. Sign up with **GitHub** (easiest - auto-connects)
4. ✅ **NO CREDIT CARD REQUIRED** for free tier

---

### Step 3: Create Web Service ⏱️ 3 minutes

1. Click **"New +"** (top right) → **"Web Service"**
2. Click **"Connect"** next to your GitHub repository
   - If you don't see it, click "Configure account" to grant access
3. Find: **"Python Engine"** repository
4. Click **"Connect"**

---

### Step 4: Configure Service ⏱️ 5 minutes

Render will auto-detect your `render.yaml`, verify these settings:

| Setting | Value |
|---------|-------|
| **Name** | `shm-reports` (or any name you want) |
| **Environment** | `Python 3` |
| **Build Command** | `./build.sh` |
| **Start Command** | `./start-render.sh` |
| **Plan** | `Free` (or `Starter` for $7/month) |
| **Branch** | `main` |

Click **"Advanced"** to verify:
- **Health Check Path:** `/api/v1/health` ✅
- **Auto-Deploy:** `Yes` ✅

---

### Step 5: Add Google Credentials ⏱️ 3 minutes

**🔥 MOST IMPORTANT STEP!**

1. Scroll to **"Environment Variables"**
2. Click **"Add Environment Variable"**
3. Set these variables:

#### Environment Variables to Add:

| Key | Value | How to Get It |
|-----|-------|---------------|
| `GOOGLE_CREDENTIALS_JSON` | *Your JSON content* | Copy entire contents from `Credentials/focal-set-486609-s9-f052f8c1a756.json` |
| `PYTHON_VERSION` | `3.11.11` | Already set in render.yaml |
| `SHM_LOG_LEVEL` | `INFO` | Already set in render.yaml |

#### How to Add Google Credentials:

1. Open your credentials file:
   ```
   Credentials/focal-set-486609-s9-f052f8c1a756.json
   ```

2. Copy the **ENTIRE** content (looks like this):
   ```json
   {
     "type": "service_account",
     "project_id": "focal-set-486609-s9",
     "private_key_id": "...",
     "private_key": "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n",
     "client_email": "...",
     ...
   }
   ```

3. In Render:
   - Click **"Add Environment Variable"**
   - Key: `GOOGLE_CREDENTIALS_JSON`
   - Value: **Paste the entire JSON** (yes, the whole thing!)
   - Click **"Save"**

---

### Step 6: Deploy! ⏱️ 5 minutes

1. Review all settings
2. Click **"Create Web Service"** (bottom of page)
3. Watch the deployment logs in real-time! 🎬

You'll see:
```
==> Cloning from GitHub...
==> Running build command: ./build.sh
==> Installing dependencies...
==> Starting service: ./start-render.sh
==> Service is live! 🎉
```

⏱️ **First deployment takes 3-5 minutes**

---

### Step 7: Get Your URL ⏱️ 1 minute

Once deployed, Render gives you a URL like:

```
https://shm-reports.onrender.com
```

This is your **public URL**! Anyone can use it.

---

### Step 8: Test Your Deployment ⏱️ 2 minutes

#### Test 1: Health Check

Open in browser or use PowerShell:

```powershell
curl https://shm-reports.onrender.com/api/v1/health
```

Expected response:
```json
{
  "status": "ok",
  "engine": "SHM Slide Layout Engine v1.0"
}
```

#### Test 2: Open Frontend

Open in browser:
```
https://shm-reports.onrender.com/
```

You should see your SHM Report Generator interface!

#### Test 3: Generate a Report

1. Select a client from dropdown
2. Choose dates
3. Click "Generate Report"
4. Wait 20-40 seconds (first time might be slower)
5. Click "Download PPTX"

✅ **Success!** Your app is live!

---

## 🎯 Quick Reference

### Your URLs (after deployment)

Replace `shm-reports` with your actual service name:

- **Frontend:** `https://shm-reports.onrender.com/`
- **Health Check:** `https://shm-reports.onrender.com/api/v1/health`
- **Generate API:** `https://shm-reports.onrender.com/api/v1/generate`
- **Download:** `https://shm-reports.onrender.com/api/v1/download/{report_id}`

### Managing Your Service

- **View Logs:** Render Dashboard → Your Service → Logs
- **Redeploy:** Render Dashboard → Manual Deploy → "Deploy latest commit"
- **Update Env Vars:** Render Dashboard → Environment → Edit
- **Upgrade Plan:** Render Dashboard → Settings → Change Plan to "Starter" ($7/month)

### Making Updates

After deployment, when you make changes:

```powershell
git add .
git commit -m "Your update message"
git push origin main
```

Render will **auto-deploy** in 2-3 minutes! ✨

---

## ⚠️ Important Notes

### Free Tier Limitations

- **Cold starts:** Service spins down after 15 minutes of inactivity
- **First request after idle:** Takes 30-60 seconds to wake up
- **Users might experience:** Initial delay on first use

### Solutions

1. **For testing/personal use:** Free tier is perfect! ✅
2. **For public use:** Upgrade to Starter ($7/month) - no cold starts
3. **Keep-alive (free):** Use cron-job.org to ping health endpoint every 10 minutes

### Security

- ✅ Credentials are in environment variables (not in code)
- ✅ `.gitignore` prevents pushing secrets to GitHub
- ✅ Render uses HTTPS automatically
- ✅ Free SSL certificate included

---

## 🆘 Troubleshooting

### Issue: Build Fails

**Check:** Render Dashboard → Logs → Look for error messages

**Common fix:** Make sure `build.sh` and `start-render.sh` have LF line endings (not CRLF)

### Issue: "Drive agent disabled"

**Check:** Environment variable `GOOGLE_CREDENTIALS_JSON` is set correctly

**Fix:**
1. Go to Environment tab
2. Verify `GOOGLE_CREDENTIALS_JSON` exists
3. Make sure it contains the full JSON (starts with `{`, ends with `}`)
4. Manual Deploy → "Clear build cache & deploy"

### Issue: 502 Bad Gateway

**Check:** Service is starting up (first deploy or cold start)

**Wait:** 30-60 seconds, then refresh

### Issue: Reports generate but download fails

**Check:** Output directory permissions

**This shouldn't happen** - `build.sh` creates the output directory

---

## 🎉 You're Done!

Your Python Engine is now deployed and accessible to anyone with the URL!

**Share your URL with users:**
```
https://shm-reports.onrender.com
```

---

## 📊 Cost Summary

| Plan | Cost/Month | Best For | Cold Starts |
|------|------------|----------|-------------|
| **Free** | $0 | Testing, personal use | Yes (15 min) |
| **Starter** | $7 | Public use, small teams | No - Always on |
| **Standard** | $25 | Production, high traffic | No + More resources |

**Recommendation:** Start with **Free**, upgrade to **Starter** when you have regular users.

---

**Questions?** Check the logs first, then review this checklist! 🚀
