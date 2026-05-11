# Deploying to Vercel with Docker

## Prerequisites
1. GitHub account with your code pushed
2. Vercel account (vercel.com)
3. Google Cloud credentials file

## Step 1: Prepare Google Credentials

You need to securely pass your Google credentials to the deployed app. Two options:

### Option A: Use Environment Variables (Recommended)
1. In `data_loader.py`, the code already checks `st.secrets["gcp_service_account"]`
2. In Vercel dashboard:
   - Go to your project Settings → Environment Variables
   - Add key: `GCP_SERVICE_ACCOUNT`
   - Value: Your entire `google_oauth_credentials.json` content (as JSON)

### Option B: Store in Vercel Secrets (for OAuth)
```
GCP_OAUTH_CLIENT_ID=your_client_id
GCP_OAUTH_CLIENT_SECRET=your_client_secret
GCP_OAUTH_REDIRECT_URI=https://your-app.vercel.app/callback
```

## Step 2: Push Code to GitHub

```bash
cd ~/fulfillment-metrics
git add Dockerfile vercel.json .dockerignore VERCEL_DEPLOYMENT.md
git commit -m "Add Docker and Vercel deployment configuration"
git push origin main
```

## Step 3: Deploy to Vercel

### Method 1: Via Vercel CLI (Recommended)
```bash
npm install -g vercel
vercel --prod
```

### Method 2: Via Vercel Dashboard
1. Go to vercel.com
2. Click "New Project"
3. Select your GitHub repository
4. Vercel auto-detects the Docker setup
5. Add environment variables in Settings
6. Click "Deploy"

## Step 4: Configure Environment Variables in Vercel

1. In Vercel Dashboard → Project Settings → Environment Variables
2. Add your Google credentials:

```
GCP_SERVICE_ACCOUNT = {"type": "service_account", "project_id": "...", ...}
```

Or for OAuth:
```
GCP_OAUTH_CLIENT_ID = your_client_id
GCP_OAUTH_CLIENT_SECRET = your_client_secret  
```

## Step 5: Modify data_loader.py for Production

The `data_loader.py` already checks for `st.secrets["gcp_service_account"]`, which will work on Vercel.

If using OAuth on Vercel (not recommended for long-running apps), you'd need additional setup for the OAuth flow.

## Important Notes

### Vercel Limitations:
- **Free tier**: Limited to 3 concurrent deployments
- **Serverless functions**: Have 60-second timeout (Streamlit needs longer)
- **Pro plan**: Better for long-running apps ($20/month)
- **Alternative**: Render or Railway may be better for Streamlit (have free tiers for long-running apps)

### Production Checklist:
- [ ] Google credentials set as environment variables
- [ ] Database connection tested
- [ ] Cache TTL verified (currently 1 hour)
- [ ] Error logging in place
- [ ] CORS enabled if needed
- [ ] Rate limiting considered

## Troubleshooting

**Build fails with "Docker not available":**
- Ensure Dockerfile exists in root directory ✓
- Verify vercel.json references Dockerfile ✓

**App times out on first load:**
- Vercel free tier has limitations
- Consider upgrading to Pro or using alternative (Render, Railway)
- Increase cache TTL to reduce API calls

**Google credentials not working:**
- Check environment variable name matches exactly
- Verify JSON is valid (use JSONValidator online)
- Check service account has required permissions

## Monitoring

After deployment:
1. Vercel Dashboard → Deployments → View deployment
2. Check logs for errors
3. Test the dashboard at your Vercel URL

## Local Testing with Docker

Before deploying, test locally:

```bash
docker build -t fulfillment-metrics .
docker run -p 8501:8501 \
  -e GCP_SERVICE_ACCOUNT='{"type":"service_account",...}' \
  fulfillment-metrics
```

Then open http://localhost:8501
