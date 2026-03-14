# Outlook → Notion Task Add-in

An Outlook add-in that lets you create Notion task board cards directly from emails, with a hyperlink back to the source email.

## Features
- One-click button in the Outlook ribbon while reading an email
- Auto-fills task title from email subject
- Attaches a clickable hyperlink back to the original email inside the Notion card
- Optionally includes email body preview as a quote block
- Supports Due Date, Priority, and custom Status values
- Settings panel for your Notion token and database ID

## Setup

### 1. Prerequisites
- Node.js (v16+) installed
- An Outlook account (Microsoft 365 / Outlook.com)
- A Notion account with a database (board/table)

### 2. Install & generate SSL cert
```bash
cd outlook-notion-addin
npm run gen-cert   # generates certs/server.key and certs/server.cert
node server.js     # starts the local HTTPS server on port 3000
```

> On macOS/Linux you may need to trust the self-signed cert:
> Open Keychain Access → import certs/server.cert → set to "Always Trust"

### 3. Create a Notion Integration
1. Go to https://www.notion.so/my-integrations
2. Click "+ New integration"
3. Give it a name (e.g. "Outlook Tasks"), select your workspace
4. Copy the **Internal Integration Token** (starts with `secret_`)
5. Open your Notion task database → click ••• → Connections → add your integration

### 4. Get your Notion Database ID
Your database URL looks like:
`https://www.notion.so/workspace/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX?v=...`
The 32-character string is your **Database ID**.

### 5. Sideload the add-in into Outlook
#### Outlook on the web (easiest):
1. Open https://outlook.office.com
2. Go to Settings → Mail → Customize actions → Manage add-ins
3. Click the ⊕ icon → "Add from file..."
4. Upload `manifest.xml`

#### Outlook Desktop (Windows):
1. Open Outlook → File → Manage Add-ins → "Add a custom add-in" → "Add from file"
2. Select `manifest.xml`

#### Outlook Desktop (Mac):
1. Open Outlook → Tools → Add-ins → Manage → Upload `manifest.xml`

### 6. Configure the add-in
1. Open any email in Outlook
2. Click **"Add to Notion"** in the ribbon
3. Click the ⚙️ Settings gear icon
4. Paste your Notion Integration Token and Database ID
5. Set the Status property name to match your database (default: "Status")
6. Click **Save Settings**

## File Structure
```
outlook-notion-addin/
├── manifest.xml       # Outlook add-in registration
├── taskpane.html      # Add-in UI
├── taskpane.css       # Styles
├── taskpane.js        # Logic + Notion API integration
├── commands.html      # Required placeholder for ribbon commands
├── server.js          # Local HTTPS proxy (required for CORS)
├── package.json       # npm config
├── README.md          # This file
└── certs/             # Auto-generated SSL certs (after npm run gen-cert)
    ├── server.key
    └── server.cert
```

## Notes
- The local server (server.js) must be running while you use the add-in
- The server acts as a proxy to bypass CORS restrictions on the Notion API
- Your Notion token is stored in browser localStorage, never sent anywhere except to your local server
