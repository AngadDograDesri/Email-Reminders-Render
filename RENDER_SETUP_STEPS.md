# Render Deployment Setup Steps

## After Creating the Blueprint

Once you deploy using the Blueprint feature, you need to manually configure a few environment variables.

### Step 1: Get the Web Service URL

1. Go to your Render Dashboard
2. Find the **email-reminders-api** web service
3. Copy its URL (it will be something like: `https://email-reminders-api.onrender.com`)

### Step 2: Get the Auto-Generated API Key

1. Still in the **email-reminders-api** web service
2. Go to the **Environment** tab
3. Find the `WEBHOOK_API_KEY` variable
4. Copy its auto-generated value

### Step 3: Configure the Cron Job

1. Go to the **email-reminders-cron** service
2. Go to the **Environment** tab
3. Add/Edit these variables:

```
WEBHOOK_API_URL = https://email-reminders-api.onrender.com
WEBHOOK_API_KEY = (paste the key from step 2)
```

### Step 4: Add Your Required Variables

Still in the cron job's Environment tab, add:

```
TENANT_ID = your-microsoft-tenant-id
CLIENT_ID = your-microsoft-client-id
CLIENT_SECRET = your-microsoft-client-secret
OPENAI_API_KEY = your-openai-api-key
USER_EMAILS = user1@domain.com,user2@domain.com
```

### Step 5: Verify Deployment

1. **Test the API**:
   - Visit: `https://email-reminders-api.onrender.com/api/health`
   - Should return: `{"status": "healthy", "database_type": "PostgreSQL", ...}`

2. **Check the Cron Job**:
   - Go to the cron job service
   - Click on "Logs" tab
   - Wait for the next scheduled run (12:00 UTC / 5:30 PM IST)
   - Verify it runs successfully

### Why These Steps Are Needed

Render's Blueprint doesn't support the `url` property in `fromService` references. The valid properties are:
- `connectionString` (for databases)
- `host`
- `hostport`
- `port`

Since we need the full HTTPS URL for the webhook API, we have to set it manually after deployment.

## Quick Reference

| Service | What to Configure |
|---------|------------------|
| **email-reminders-api** | Nothing (auto-configured) |
| **email-reminders-cron** | WEBHOOK_API_URL, WEBHOOK_API_KEY, TENANT_ID, CLIENT_ID, CLIENT_SECRET, OPENAI_API_KEY, USER_EMAILS |
| **email-reminders-db** | Nothing (auto-configured) |

## Troubleshooting

**If the cron job can't connect to the API:**
- Verify `WEBHOOK_API_URL` is set correctly (with https://)
- Verify `WEBHOOK_API_KEY` matches the one in the web service
- Check that `USE_EXCLUSION_API` is set to `true`

**If authentication fails:**
- Double-check the `WEBHOOK_API_KEY` value matches exactly
- No extra spaces or quotes

**If the API is not responding:**
- Free tier web services spin down after inactivity
- First request will take ~30 seconds to wake up
- The cron job will automatically wait for the service to wake

## Complete Environment Variables List

### Web Service (email-reminders-api)
✅ Auto-configured by render.yaml:
- `DATABASE_URL`
- `WEBHOOK_HOST`
- `WEBHOOK_PORT`
- `AUTO_CLEANUP_DAYS`
- `WEBHOOK_API_KEY`

### Cron Job (email-reminders-cron)
✅ Auto-configured by render.yaml:
- `DATABASE_URL`
- `USE_EXCLUSION_API`

⚠️ Must set manually:
- `WEBHOOK_API_URL` (from web service URL)
- `WEBHOOK_API_KEY` (copy from web service)
- `TENANT_ID` (your Microsoft Azure tenant ID)
- `CLIENT_ID` (your Microsoft Azure client ID)
- `CLIENT_SECRET` (your Microsoft Azure secret)
- `OPENAI_API_KEY` (your OpenAI key)
- `USER_EMAILS` (comma-separated list)

## Schedule

The cron job runs at:
- **5:30 PM IST** (India Standard Time)
- **12:00 UTC** (Coordinated Universal Time)
- **7:00 AM EST** (Eastern Standard Time)

Cron expression: `0 12 * * *` (every day at 12:00 UTC)

