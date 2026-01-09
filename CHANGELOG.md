# Changelog - PostgreSQL Migration & Render Deployment

## Changes Made

### 1. Database Migration: SQLite → PostgreSQL
- Created `db_utils.py` - Universal database utility module
  - Automatically detects environment (PostgreSQL vs SQLite)
  - Provides unified interface for both databases
  - Handles parameter placeholder conversion (? → %s)
  
### 2. Updated All Database Access
- `exclusion_checker.py` - Now uses db_utils
- `init_exclusions_db.py` - Now uses db_utils
- `mark_dealt_with_api.py` - Now uses db_utils
- All files work with both PostgreSQL (production) and SQLite (local dev)

### 3. Render Configuration (`render.yml`)
- **Database**: PostgreSQL free tier
  - Database name: `email_reminders`
  - Auto-configured connection string
  
- **Web Service**: Flask API
  - Endpoint: `/api/health`, `/api/mark-dealt-with`, etc.
  - Auto-configured DATABASE_URL
  - Health check enabled
  
- **Cron Job**: Daily email reminders
  - Schedule: **5:30 PM IST** (12:00 UTC)
  - Cron expression: `0 12 * * *`
  - Auto-configured to use API and database

### 4. Dependencies Updated (`requirements.txt`)
- Added: `psycopg2-binary>=2.9.9` for PostgreSQL
- Added: `gunicorn>=21.2.0` for production WSGI server
- Updated: All dependencies verified for Render deployment

### 5. Git Configuration Fixed
- Changed remote URL from SSH to HTTPS
- Resolved authentication conflict between GitHub accounts
- Successfully pushed to `AngadDograDesri/Email-Reminders-Render`

### 6. Documentation Added
- `README_RENDER_DEPLOYMENT.md` - Complete deployment guide
- `ENV_VARIABLES.md` - Environment variables reference
- `CHANGELOG.md` - This file

## Time Zone Configuration

**Schedule**: 5:30 PM IST Daily
- IST: 17:30 (UTC+5:30)
- UTC: 12:00
- Cron: `0 12 * * *`

## Architecture

```
┌─────────────────┐
│   PostgreSQL    │  Free tier database
│   Database      │  (email_reminders)
└────────┬────────┘
         │
         ├──────────┐
         │          │
┌────────▼────┐  ┌──▼─────────────┐
│  Flask API  │  │   Cron Job     │
│  (Web)      │  │   (Daily)      │
│  Port 10000 │  │   12:00 UTC    │
└─────────────┘  └────────────────┘
```

## Deployment Steps

1. **Push to GitHub** ✅
   ```bash
   git push origin main
   ```

2. **Deploy on Render**
   - Use Blueprint option
   - Connect GitHub repo
   - Render auto-configures from `render.yml`

3. **Add Environment Variables**
   - TENANT_ID
   - CLIENT_ID
   - CLIENT_SECRET
   - OPENAI_API_KEY
   - USER_EMAILS

4. **Verify**
   - Check `/api/health` endpoint
   - Monitor cron job logs
   - Test "mark as dealt with" feature

## Features Preserved

✅ Mark emails as "dealt with"  
✅ Auto-cleanup after 14 days  
✅ Multi-user support  
✅ API authentication with key  
✅ Health check endpoint  
✅ Beautiful HTML responses  

## New Features

✨ PostgreSQL support (production-ready)  
✨ Automatic database detection  
✨ Render Blueprint deployment  
✨ Scheduled cron job (5:30 PM IST)  
✨ Environment-aware configuration  
✨ Improved error handling  

## Breaking Changes

⚠️ None - backward compatible with local SQLite development

## Testing

- ✅ No linter errors
- ✅ Database utilities tested
- ✅ API endpoints verified
- 🔄 Deployment testing pending (after Render setup)

## Next Steps

1. Deploy to Render using Blueprint
2. Add environment variables in Render dashboard
3. Test API health endpoint
4. Verify cron job execution at scheduled time
5. Monitor logs for any issues

## Rollback Plan

If issues occur, revert to SQLite:
1. Don't set `DATABASE_URL` environment variable
2. Application will automatically use SQLite
3. All features work identically with local DB file

