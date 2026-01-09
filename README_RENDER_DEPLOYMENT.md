# Email Reminders - Render Deployment Guide

This project is configured to deploy on Render using the Blueprint feature with PostgreSQL database.

## Architecture

- **PostgreSQL Database**: Stores excluded email instances
- **Web Service (API)**: Flask API for marking emails as dealt with
- **Cron Job**: Runs daily at 5:30 PM IST (12:00 UTC) to send email reminders

## Deployment Steps

### 1. Push to GitHub

Make sure all your code is pushed to GitHub:

```bash
git add .
git commit -m "Configure for Render deployment with PostgreSQL"
git push origin main
```

### 2. Deploy to Render

1. Go to [Render Dashboard](https://dashboard.render.com/)
2. Click **"New"** → **"Blueprint"**
3. Connect your GitHub repository: `AngadDograDesri/Email-Reminders-Render`
4. Render will automatically detect `render.yml` and configure:
   - PostgreSQL database (`email-reminders-db`)
   - Web service (`email-reminders-api`)
   - Cron job (`email-reminders-cron`)

### 3. Configure Environment Variables

You need to add your email service credentials. Go to each service and add:

#### For the Cron Job (`email-reminders-cron`):
- `TENANT_ID`: Your Microsoft Azure tenant ID
- `CLIENT_ID`: Your Microsoft Azure app client ID
- `CLIENT_SECRET`: Your Microsoft Azure app secret
- `OPENAI_API_KEY`: Your OpenAI API key
- Any other user-specific variables from your `.env` file

#### For the Web Service (`email-reminders-api`):
- `WEBHOOK_API_KEY`: Auto-generated (already configured in render.yml)
- `DATABASE_URL`: Auto-configured from database

### 4. Verify Deployment

#### Check the Web Service:
Visit: `https://email-reminders-api.onrender.com/api/health`

Should return:
```json
{
  "status": "healthy",
  "database_type": "PostgreSQL",
  "total_exclusions": 0
}
```

#### Check the Cron Job:
- Go to the cron job service in Render
- Check the logs after the scheduled time (12:00 UTC / 5:30 PM IST)
- Verify emails are being sent

## Schedule Configuration

The cron job is configured to run at:
- **5:30 PM IST** (India Standard Time)
- **12:00 UTC** (Coordinated Universal Time)

Cron expression: `0 12 * * *`

To change the schedule, modify the `schedule` field in `render.yml`:
```yaml
schedule: "0 12 * * *"  # minute hour day month weekday
```

## Database

The application automatically detects whether to use PostgreSQL or SQLite:
- **Production (Render)**: Uses PostgreSQL (via `DATABASE_URL` env var)
- **Local Development**: Uses SQLite (`excluded_instances.db`)

### Database Schema

```sql
CREATE TABLE excluded_instances (
    id SERIAL PRIMARY KEY,
    conversation_id TEXT NOT NULL,
    latest_message_id TEXT NOT NULL,
    user_email TEXT NOT NULL,
    excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    reason TEXT,
    UNIQUE(conversation_id, latest_message_id, user_email)
);
```

## API Endpoints

The web service provides these endpoints:

- `POST/GET /api/mark-dealt-with` - Mark email as dealt with
- `GET /api/check-excluded/<conversationId>/<latestMessageId>/<userEmail>` - Check if excluded
- `GET /api/exclusions/<userEmail>` - List all exclusions for user
- `POST/GET /api/undo-exclusion` - Remove an exclusion
- `GET /api/health` - Health check

## Auto-Cleanup

The system automatically deletes exclusions older than 14 days. Configure this with:
```yaml
envVars:
  - key: AUTO_CLEANUP_DAYS
    value: 14
```

## Local Development

To run locally with SQLite:

```bash
# Install dependencies
pip install -r requirements.txt

# Initialize database
python init_exclusions_db.py

# Start API server
python mark_dealt_with_api.py

# Run email check (in another terminal)
python email_followup_graph_multi_user_v2.py
```

## Troubleshooting

### Cron job not running
- Check the cron job logs in Render dashboard
- Verify all environment variables are set
- Ensure the schedule is in UTC time

### Database connection errors
- Check that `DATABASE_URL` is properly configured
- Verify the database service is running
- Check service logs for connection errors

### API not responding
- Check web service logs
- Verify the service is running (not sleeping on free tier)
- Test health endpoint: `/api/health`

## Migration from SQLite

If you have existing SQLite data, you'll need to migrate it manually:

1. Export data from SQLite:
```python
import sqlite3
conn = sqlite3.connect('excluded_instances.db')
cursor = conn.cursor()
cursor.execute("SELECT * FROM excluded_instances")
data = cursor.fetchall()
```

2. Import to PostgreSQL (use the API endpoints or direct SQL)

## Cost

On Render's free tier:
- **PostgreSQL**: Free (90 days, then expires)
- **Web Service**: Free (spins down after inactivity)
- **Cron Job**: Free (750 hours/month)

For production, upgrade to paid plans for persistent database and no sleep times.

## Support

For issues:
1. Check service logs in Render dashboard
2. Verify environment variables
3. Test API health endpoint
4. Check database connection

## References

- [Render Blueprints](https://render.com/docs/blueprint-spec)
- [Render Cron Jobs](https://render.com/docs/cronjobs)
- [Render PostgreSQL](https://render.com/docs/databases)

