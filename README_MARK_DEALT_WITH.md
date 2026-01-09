# Email Reminders - Mark as Dealt With Feature

This feature allows users to mark specific email instances as "dealt with" from the email digest, preventing them from appearing in future analyses until new messages arrive in the same thread.

## 🚀 Production Deployment (Render)

This project is configured for deployment on Render with:
- **PostgreSQL Database** (auto-configured)
- **Flask API Web Service** (mark as dealt with endpoints)
- **Cron Job** running at **5:30 PM IST (12:00 UTC) daily**

### Deploy to Render
1. Push code to GitHub
2. In Render Dashboard: New → Blueprint
3. Connect repository: `AngadDograDesri/Email-Reminders-Render`
4. Render auto-configures from `render.yml`
5. Add required environment variables:
   - `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` (Microsoft Azure)
   - `OPENAI_API_KEY` (OpenAI)
   - `USER_EMAILS` (comma-separated email addresses)

### Verify Deployment
```bash
curl https://your-service.onrender.com/api/health
```

**See `render.yml` for complete deployment configuration.**

## Architecture

The feature consists of:

1. **`mark_dealt_with_api.py`** - Flask API server (stores/checks exclusions)
2. **`exclusion_checker.py`** - Helper module for analysis script
3. **`init_exclusions_db.py`** - Database initialization
4. **`db_utils.py`** - Database abstraction (PostgreSQL/SQLite)

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Initialize Database

Run once to create the database:

```bash
python init_exclusions_db.py
```

**Production**: Automatically uses PostgreSQL (via `DATABASE_URL` env var)  
**Local Dev**: Uses SQLite (`excluded_instances.db`)

### 3. Configure Environment Variables

**Production (Render)**: Most variables are auto-configured in `render.yml`

**Local Development** - Add to your `.env` file:

```env
# Microsoft Azure (Required)
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret

# OpenAI (Required)
OPENAI_API_KEY=your-openai-key

# Users (Required)
USER_EMAILS=user1@domain.com,user2@domain.com

# Local Database (Optional - auto-detected)
EXCLUSIONS_DB_PATH=excluded_instances.db

# API Configuration (Optional - for local testing)
WEBHOOK_API_URL=http://localhost:5000
WEBHOOK_HOST=127.0.0.1
WEBHOOK_PORT=5000
USE_EXCLUSION_API=false
```

### 4. Start API Server (Optional)

If using API mode (`USE_EXCLUSION_API=true`):

```bash
python mark_dealt_with_api.py
```

The server will run on `http://localhost:5000` by default.

## Usage

### Mode 1: Direct Database Access (Default)

The analysis script can use `exclusion_checker.py` to check exclusions directly from the database:

```python
from exclusion_checker import is_email_instance_excluded

# In your analysis loop
if is_email_instance_excluded(conversation_id, latest_message_id, user_email):
    print("  Skipping: Marked as 'dealt with'")
    continue
```

**Pros:**
- No separate server needed
- Faster (no network calls)
- Simpler setup

**Cons:**
- Requires database file to be accessible to analysis script

### Mode 2: API Mode

Set `USE_EXCLUSION_API=true` in `.env` and start the API server. The `exclusion_checker.py` module will automatically use the API.

**Pros:**
- Centralized database
- Can be deployed separately
- Better for multi-machine setups

**Cons:**
- Requires API server to be running
- Network dependency

## API Endpoints

When running the API server, these endpoints are available:

### Mark as Dealt With
```
POST /api/mark-dealt-with
GET  /api/mark-dealt-with?conversationId=...&latestMessageId=...&userEmail=...

Body (POST):
{
  "conversationId": "...",
  "latestMessageId": "...",
  "userEmail": "...",
  "reason": "..."  // optional
}
```

### Check if Excluded
```
GET /api/check-excluded/<conversationId>/<latestMessageId>/<userEmail>
```

### List User Exclusions
```
GET /api/exclusions/<userEmail>
```

### Undo Exclusion
```
POST /api/undo-exclusion
GET  /api/undo-exclusion?conversationId=...&latestMessageId=...&userEmail=...
```

### Health Check
```
GET /api/health
```

## Integration with Analysis Script

When ready to integrate, modify `email_followup_graph_multi_user_v2.py`:

1. Import the checker:
```python
from exclusion_checker import is_email_instance_excluded
```

2. Add check after getting latest message:
```python
latest_message = get_latest_message_in_conversation(conversation_messages, user_email)
latest_message_id = latest_message.get("id", "")

if latest_message_id and is_email_instance_excluded(conversation_id, latest_message_id, user_email):
    print(f"  Skipping: This email instance was marked as 'dealt with'")
    skipped_count += 1
    continue
```

3. Include `latest_message_id` in email data for HTML generation:
```python
email_data = {
    'conversation_id': conversation_id,
    'latest_message_id': latest_message_id,  # Add this
    'user_email': user_email,  # Add this
    # ... other fields
}
```

4. Add button to HTML in `build_section_table()`:
```python
# Add button column
mark_button = f'''
<a href="{WEBHOOK_API_URL}/api/mark-dealt-with?conversationId={conversation_id}&latestMessageId={latest_message_id}&userEmail={user_email}" 
   style="background-color:#4caf50; color:white; padding:6px 12px; text-decoration:none; border-radius:4px; font-size:12px;">
   ✓ Mark as Dealt With
</a>
'''
```

## How It Works

- **Instance-Based Exclusion**: Each exclusion is tied to a specific `conversationId` + `latestMessageId` + `userEmail` combination
- **Thread Reappearance**: If a new message arrives in the same thread (new `latestMessageId`), the conversation will reappear in future digests
- **Persistence**: Exclusions are stored in SQLite database and persist across script runs

## Example Flow

1. **Day 1**: Email digest shows conversation "Project Update" with latest message ID `msg-123`
2. **User clicks**: "Mark as Dealt With" button
3. **Database stores**: `conversationId="conv-abc"`, `latestMessageId="msg-123"`, `userEmail="peter@desri.com"`
4. **Day 2**: Same conversation still has `msg-123` → Analysis script skips it (marked as dealt with)
5. **Day 3**: New message arrives, latest message ID is now `msg-456` → Conversation reappears in digest (new instance)

## Testing

1. Initialize database: `python init_exclusions_db.py`
2. Start API server: `python mark_dealt_with_api.py`
3. Test API: `curl http://localhost:5000/api/health`
4. Mark exclusion: `curl "http://localhost:5000/api/mark-dealt-with?conversationId=test&latestMessageId=msg1&userEmail=test@example.com"`
5. Check exclusion: `curl http://localhost:5000/api/check-excluded/test/msg1/test@example.com`