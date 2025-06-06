# Google Cloud Console Setup Instructions

## Step-by-Step Setup

### 1. Go to Google Cloud Console
Visit: https://console.cloud.google.com/

### 2. Create or Select a Project
- If you don't have a project, click "Select a project" → "New Project"
- Enter a project name (e.g., "MCP Google Docs")
- Click "Create"

### 3. Enable Required APIs
- Go to "APIs & Services" → "Library"
- Search for and enable:
  - **Google Docs API**
  - **Google Drive API**

### 4. Create OAuth 2.0 Credentials
- Go to "APIs & Services" → "Credentials"
- Click "Create Credentials" → "OAuth client ID"
- If prompted, configure the OAuth consent screen:
  - Choose "External" (unless you have a Google Workspace account)
  - Fill in required fields:
    - App name: "MCP Google Docs"
    - User support email: your email
    - Developer contact: your email
  - Add scopes if prompted:
    - `https://www.googleapis.com/auth/documents`
    - `https://www.googleapis.com/auth/drive`
- For Application type, select **"Desktop app"**
- Give it a name (e.g., "MCP Google Docs Client")
- Click "Create"

### 5. Download Credentials
- After creation, you'll see your client ID and secret
- Click "Download JSON" button
- Save the downloaded file as `credentials.json` in your project root directory
- **IMPORTANT**: Replace the `credentials.json.template` file with your actual `credentials.json`

### 6. Test the Setup
```bash
npm start
```

The first time you run this, it will:
1. Open a browser window for Google OAuth
2. Ask you to sign in and authorize the app
3. Create a `token.json` file with your access tokens

## Security Notes
- Never commit `credentials.json` or `token.json` to version control
- These files contain sensitive authentication information
- The `.gitignore` file already excludes these files

## Troubleshooting
- If you get "access denied" errors, make sure the APIs are enabled
- If authentication fails, delete `token.json` and try again
- Make sure your Google account has access to Google Drive and Docs 