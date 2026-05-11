# Complete Deployment Setup Guide

This project is configured to deploy to Vercel automatically through GitHub Actions. Follow these steps to complete the deployment setup.

## Step 1: Create a Vercel Account (if you don't have one)

1. Visit https://vercel.com/signup
2. Sign up using your GitHub account
3. Complete the onboarding process

## Step 2: Create a Vercel Project

### Option A: Manual Setup (Recommended if automatic linking doesn't work)

1. Go to https://vercel.com/dashboard
2. Click "New Project"
3. Import the `fulfillment-metrics` repository from GitHub
4. Vercel will auto-detect the Docker configuration
5. Note your:
   - Organization ID (in Settings → General, usually after `/` in the URL)
   - Project ID (displayed on the project dashboard)

### Option B: Using the GitHub App

1. Go to https://github.com/apps/vercel
2. Click "Install"
3. Select the `fulfillment-metrics` repository
4. Complete the connection

## Step 3: Get Your Vercel Credentials

After creating your project, you need:

- **VERCEL_TOKEN**: Personal API Token for authentication
  1. Go to https://vercel.com/account/tokens
  2. Click "Create" 
  3. Name it "fulfillment-metrics-deploy"
  4. Copy the token

- **VERCEL_ORG_ID**: Your Organization ID
  1. Go to https://vercel.com/account/settings
  2. Find "Organization ID" on the General tab
  3. Copy it

- **VERCEL_PROJECT_ID**: Your Project ID
  1. Go to https://vercel.com/dashboard
  2. Select the fulfillment-metrics project
  3. Go to Settings → General
  4. Find "Project ID" and copy it

## Step 4: Add GitHub Secrets

1. Go to GitHub: https://github.com/Danaraai/fulfillment-metrics
2. Click "Settings" (top right of repo)
3. Go to "Secrets and variables" → "Actions"
4. Click "New repository secret"
5. Add these secrets:
   - Name: `VERCEL_TOKEN` → Value: (paste your token)
   - Name: `VERCEL_ORG_ID` → Value: (paste your org ID)
   - Name: `VERCEL_PROJECT_ID` → Value: (paste your project ID)

## Step 5: Configure Environment Variables in Vercel

After the first deployment, you need to set up the Google credentials:

1. Go to your Vercel project: https://vercel.com/dashboard
2. Click on "fulfillment-metrics"
3. Go to Settings → Environment Variables
4. Add:
   - Key: `GCP_SERVICE_ACCOUNT`
   - Value: Paste your entire `google_oauth_credentials.json` content (as JSON)
5. Make sure to select all environments (Production, Preview, Development)

## Step 6: Trigger the Deployment

The GitHub Action will automatically trigger on the next push to `main`. To test it:

```bash
cd /Users/danara/fulfillment-metrics
git pull origin main
# Make a small change or just push current state
git push origin main
```

This will trigger the GitHub Action, which will deploy to Vercel.

## Verification

After deployment:

1. Check GitHub Actions: https://github.com/Danaraai/fulfillment-metrics/actions
2. Find the "Deploy to Vercel" workflow and click it
3. Check the deployment log
4. Visit your Vercel project URL (format: `fulfillment-metrics-xxx.vercel.app`)

## Troubleshooting

### "Credentials not found" error
- Check that GCP_SERVICE_ACCOUNT is set in Vercel project settings
- Verify the JSON is valid and complete

### "Permission denied" errors
- Ensure your Google service account has permissions for:
  - Google Sheets API
  - Google Drive API
  - Any other APIs your app uses

### "Build failed" errors
- Check the Vercel deployment logs
- Ensure all Python dependencies are in requirements.txt
- Verify Dockerfile is correct

### Manual Deployment as Fallback

If automated deployment fails, deploy manually:

1. Install Vercel CLI: `npm install -g vercel` (requires Node.js)
2. Run: `vercel --prod`
3. Follow the prompts to select your project
4. Set environment variables when prompted

## Environment Variables Summary

| Variable | Purpose | Where to Get |
|----------|---------|--------------|
| `GCP_SERVICE_ACCOUNT` | Google Cloud credentials | Download from Google Cloud Console |
| `STREAMLIT_SERVER_HEADLESS` | Run Streamlit in server mode | Already set in vercel.json |
| `STREAMLIT_SERVER_PORT` | Port for Streamlit | Already set in vercel.json (8501) |

## Important Notes

- The GitHub Actions workflow runs automatically on every push to `main`
- Initial deployment may take 5-10 minutes
- Subsequent deployments are faster due to caching
- Monitor the GitHub Actions logs for any issues
- Check Vercel's dashboard for real-time deployment status

## Next Steps After Successful Deployment

1. Test your dashboard at the Vercel URL
2. Set up custom domain (optional)
3. Configure CI/CD settings in Vercel (optional)
4. Monitor application logs in Vercel dashboard
