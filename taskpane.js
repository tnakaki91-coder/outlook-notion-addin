/* ============================================================
   Outlook -> Notion Task Add-in
   taskpane.js
   ============================================================ */

'use strict';

const STORAGE_KEYS = {
  TOKEN: 'notion_token',
  DB_ID: 'notion_db_id',
  STATUS_PROP: 'notion_status_prop',
  DEFAULT_STATUS: 'notion_default_status',
  ASSIGNEE_PROP: 'notion_assignee_prop'
};

let currentEmail = null;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  bindUI();
  loadSettings();
  const configured = isConfigured();
  if (!configured) {
    show('not-configured');
    hide('task-form');
    hide('email-preview');
  } else {
    hide('not-configured');
    await loadEmailData();
    show('email-preview');
    show('task-form');
  }
});

function show(id)  { document.getElementById(id)?.classList.remove('hidden'); }
function hide(id)  { document.getElementById(id)?.classList.add('hidden'); }
function get(id)   { return document.getElementById(id); }

function showError(msg) {
  get('error-text').textContent = msg;
  show('error-msg');
  hide('success-msg');
}

function setLoading(loading) {
  const btn = get('create-task-btn');
  btn.disabled = loading;
  loading ? hide('btn-label') : show('btn-label');
  loading ? show('btn-spinner') : hide('btn-spinner');
}

function bindUI() {
  get('settings-btn').addEventListener('click', () => {
    show('settings-view');
    hide('main-view');
  });
  get('back-btn').addEventListener('click', () => {
    hide('settings-view');
    show('main-view');
  });
  get('go-to-settings-btn')?.addEventListener('click', () => {
    show('settings-view');
    hide('main-view');
  });
  get('save-settings-btn').addEventListener('click', saveSettings);
  get('create-task-btn').addEventListener('click', createNotionTask);
  get('create-another-btn').addEventListener('click', () => {
    hide('success-msg');
    show('task-form');
    show('email-preview');
    get('task-description').value = '';
    get('task-due-date').value = '';
    get('task-priority').value = '';
  });
  get('dismiss-error-btn').addEventListener('click', () => hide('error-msg'));
}

function isConfigured() {
  return !!(localStorage.getItem(STORAGE_KEYS.TOKEN) && localStorage.getItem(STORAGE_KEYS.DB_ID));
}

function loadSettings() {
  get('notion-token').value      = localStorage.getItem(STORAGE_KEYS.TOKEN) || '';
  get('notion-db-id').value      = localStorage.getItem(STORAGE_KEYS.DB_ID) || '';
  get('status-property').value   = localStorage.getItem(STORAGE_KEYS.STATUS_PROP) || 'Status';
  get('default-status').value    = localStorage.getItem(STORAGE_KEYS.DEFAULT_STATUS) || 'To Do';
  get('assignee-property').value = localStorage.getItem(STORAGE_KEYS.ASSIGNEE_PROP) || '';
}

function saveSettings() {
  const token   = get('notion-token').value.trim();
  const dbId    = get('notion-db-id').value.trim();
  const statusP = get('status-property').value.trim() || 'Status';
  const defStat = get('default-status').value.trim() || 'To Do';
  const assP    = get('assignee-property').value.trim();
  if (!token || !dbId) {
    showSettingsStatus('Please enter both the Integration Token and Database ID.', 'error');
    return;
  }
  localStorage.setItem(STORAGE_KEYS.TOKEN, token);
  localStorage.setItem(STORAGE_KEYS.DB_ID, dbId);
  localStorage.setItem(STORAGE_KEYS.STATUS_PROP, statusP);
  localStorage.setItem(STORAGE_KEYS.DEFAULT_STATUS, defStat);
  localStorage.setItem(STORAGE_KEYS.ASSIGNEE_PROP, assP);
  showSettingsStatus('Settings saved!', 'success');
  setTimeout(() => {
    hide('settings-view');
    show('main-view');
    hide('not-configured');
    show('email-preview');
    show('task-form');
    loadEmailData();
  }, 1000);
}

function showSettingsStatus(msg, type) {
  const el = get('settings-status');
  el.textContent = msg;
  el.className = 'status-msg ' + type;
  show('settings-status');
}

async function loadEmailData() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    const emailData = {
      subject: item.subject || '(No Subject)',
      from: '',
      fromEmail: '',
      body: '',
      itemId: item.itemId,
      webLink: item.webLink || ''
    };
    if (item.from) {
      emailData.from      = item.from.displayName || '';
      emailData.fromEmail = item.from.emailAddress || '';
    }
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: emailData }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        result.asyncContext.body = result.value.substring(0, 1000).trim();
      }
      currentEmail = result.asyncContext;
      get('preview-from').textContent    = currentEmail.from
        ? 'From: ' + currentEmail.from + ' <' + currentEmail.fromEmail + '>'
        : 'From: Unknown';
      get('preview-subject').textContent = currentEmail.subject;
      get('task-title').value = currentEmail.subject;
      resolve(currentEmail);
    });
  });
}

function buildEmailLink(emailData) {
  if (emailData.webLink && emailData.webLink.startsWith('http')) {
    return emailData.webLink;
  }
  const encoded = encodeURIComponent(emailData.itemId || '');
  return 'https://outlook.office.com/mail/deeplink/compose?ItemID=' + encoded;
}

async function createNotionTask() {
  if (!currentEmail) {
    showError('Could not read email data. Please close and reopen the panel.');
    return;
  }
  const title       = get('task-title').value.trim();
  const description = get('task-description').value.trim();
  const dueDate     = get('task-due-date').value;
  const priority    = get('task-priority').value;
  const includeLink = get('include-email-link').checked;
  const includeBody = get('include-body-preview').checked;
  if (!title) { showError('Please enter a task title.'); return; }

  const token      = localStorage.getItem(STORAGE_KEYS.TOKEN);
  const dbId       = localStorage.getItem(STORAGE_KEYS.DB_ID);
  const statusProp = localStorage.getItem(STORAGE_KEYS.STATUS_PROP) || 'Status';
  const defStatus  = localStorage.getItem(STORAGE_KEYS.DEFAULT_STATUS) || 'To Do';

  if (!token || !dbId) {
    showError('Please configure your Notion settings first.');
    show('settings-view'); hide('main-view');
    return;
  }

  setLoading(true);
  hide('error-msg');

  try {
    const emailLink = buildEmailLink(currentEmail);
    const pageUrl   = await callNotionAPI(token, dbId, {
      title, description, dueDate, priority,
      includeLink, includeBody, emailLink,
      statusProp, defStatus
    });
    setLoading(false);
    hide('task-form');
    hide('email-preview');
    get('notion-link').href = pageUrl;
    show('success-msg');
  } catch (err) {
    setLoading(false);
    showError(err.message || 'Failed to create task. Check your settings and try again.');
  }
}

async function callNotionAPI(token, dbId, opts) {
  const { title, description, dueDate, priority,
          includeLink, includeBody, emailLink,
          statusProp, defStatus } = opts;

  const properties = {
    "Name": { "title": [{ "text": { "content": title } }] },
    [statusProp]: { "select": { "name": defStatus } }
  };
  if (dueDate)   properties["Due Date"]  = { "date": { "start": dueDate } };
  if (priority)  properties["Priority"]  = { "select": { "name": priority } };

  const children = [];

  if (includeLink) {
    children.push({
      "object": "block", "type": "callout",
      "callout": {
        "rich_text": [
          { "type": "text", "text": { "content": "Email Source: " } },
          { "type": "text",
            "text": { "content": currentEmail.subject, "link": { "url": emailLink } },
            "annotations": { "bold": true, "color": "blue" }
          }
        ],
        "icon": { "emoji": "Email" },
        "color": "blue_background"
      }
    });
    children.push({
      "object": "block", "type": "paragraph",
      "paragraph": {
        "rich_text": [{
          "type": "text",
          "text": { "content": "From: " + currentEmail.from + " <" + currentEmail.fromEmail + ">" },
          "annotations": { "color": "gray" }
        }]
      }
    });
  }

  if (includeLink || includeBody) {
    children.push({ "object": "block", "type": "divider", "divider": {} });
  }

  if (description) {
    children.push({
      "object": "block", "type": "heading_3",
      "heading_3": { "rich_text": [{ "type": "text", "text": { "content": "Notes" } }] }
    });
    children.push({
      "object": "block", "type": "paragraph",
      "paragraph": { "rich_text": [{ "type": "text", "text": { "content": description } }] }
    });
  }

  if (includeBody && currentEmail.body) {
    children.push({
      "object": "block", "type": "heading_3",
      "heading_3": { "rich_text": [{ "type": "text", "text": { "content": "Email Preview" } }] }
    });
    children.push({
      "object": "block", "type": "quote",
      "quote": { "rich_text": [{ "type": "text", "text": { "content": currentEmail.body.substring(0, 2000) } }] }
    });
  }

  const body = { parent: { database_id: dbId }, properties, children };

  const response = await fetch('https://localhost:3000/notion-proxy', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token, payload: body })
  });

  if (!response.ok) {
    const errBody = await response.json().catch(() => ({}));
    throw new Error(errBody.message || 'Notion API error: ' + response.status);
  }

  const data = await response.json();
  return data.url;
}