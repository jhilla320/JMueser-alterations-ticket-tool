# Google Drive Backend Setup (Render)

This app uses a small Node backend to handle Google OAuth and upload files to each user's Drive root.

## 1) Create a Google Cloud Project

1. Go to Google Cloud Console.
2. Create a new project.
3. Enable **Google Drive API**.
4. Configure the **OAuth consent screen**.
   - User type: External
   - Add yourself (and any testers) as test users until you publish.

## 2) Create OAuth Credentials

Create **OAuth Client ID** (Web application) and set:

- Authorized JavaScript origins: `https://jhilla320.github.io`
- Authorized redirect URI: `https://<your-render-app>.onrender.com/auth/callback`

Copy the Client ID and Client Secret.

## 3) Deploy Backend to Render

1. Create a new **Web Service** from this repo.
2. Build command: `npm install`
3. Start command: `npm start`
4. Add environment variables:

```
GOOGLE_CLIENT_ID=...
GOOGLE_CLIENT_SECRET=...
GOOGLE_REDIRECT_URI=https://<your-render-app>.onrender.com/auth/callback
APP_URL=https://jhilla320.github.io/JMueser-alterations-ticket-tool/
SQLITE_PATH=./data.sqlite
```

Note: Render free instances can sleep and their disk is not guaranteed long-term. For production, add a persistent disk or use a small managed database.

## 4) Frontend Configuration

Set the backend URL in `app.js` (see `DRIVE_API_BASE`).

## 5) Permissions

The app uses `drive.file` scope, which only allows files created by the app.

