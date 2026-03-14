/* ============================================================
   Outlook -> Notion Task Add-in  |  taskpane.js
   Calls Notion API directly via a CORS proxy (no local server needed)
   ============================================================ */
'use strict';

const STORAGE_KEYS = {
  TOKEN:          'notion_token',
  DB_ID:          'notion_db_id',
  STATUS_PROP:    'notion_status_prop',
  DEFAULT_STATUS: 'notion_default_status',
};

// Public CORS proxy — forwards requests to api.notion.com
// You can self-host this or use a Cloudflare Worker for production
const CORS_PROXY = 'https://corsproxy.io/?';
const NOTION_API = 'https://api.notion.com/v1/pages';

let currentEmail = null;

// ── Init ──────────────────────────────────────────────────────
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  bindUI();
  loadSettings();

  // Check if opened in quick-add mode
  const isQuickAdd = new URLSearchParams(window.location.search).get('quickadd') === 'true';

  if (!isConfigured()) {
    show('not-configured');
    hide('task-form');
    hide('email-preview');
  } else {
    hide('not-configured');
    await loadEmailData();
    show('email-preview');
    show('task-form');
    if (isQuickAdd) {
      // Auto-submit immediately with defaults
      setTimeout(() => createNotionTask(true), 500);
    }
  }
});

// ── Helpers ───────────────────────────────────────────────────
function show(id) { document.getElementById(id)?.classList.remove('hidden'); }
function hide(id) { document.getElementById(id)?.classList.add('hidden'); }
function get(id)  { return document.getElementById(id); }

function showError(msg) {
  get('error-text').textContent = msg;
  show('error-msg');
  hide('success-msg');
}

function setLoading(on) {
  const btn = get('create-task-btn');
  btn.disabled = on;
  on ? hide('btn-label') : show('btn-label');
  on ? show('btn-spinner') : hide('btn-spinner');
}

// ── UI Bindings ───────────────────────────────────────────────
function bindUI() {
  get('settings-btn').addEventListener('click', () => { show('settings-view'); hide('main-view'); });
  get('back-btn').addEventListener('click',     () => { hide('settings-view'); show('main-view'); });
  get('go-to-settings-btn')?.addEventListener('click', () => { show('settings-view'); hide('main-view'); });
  get('save-settings-btn').addEventListener('click', saveSettings);
  get('create-task-btn').addEventListener('click', () => createNotionTask(false));
  get('create-another-btn').addEventListener('click', () => {
    hide('success-msg');
    show('task-form');
    show('email-preview');
    get('task-description').value = '';
    get('task-due-date').value    = '';
    get('task-priority').value    = '';
  });
  get('dismiss-error-btn').addEventListener('click', () => hide('error-msg'));
}

// ── Settings ──────────────────────────────────────────────────
function isConfigured() {
  return !!(localStorage.getItem(STORAGE_KEYS.TOKEN) && localStorage.getItem(STORAGE_KEYS.DB_ID));
}

function loadSettings() {
  get('notion-token').value    = localStorage.getItem(STORAGE_KEYS.TOKEN)          || '';
  get('notion-db-id').value    = localStorage.getItem(STORAGE_KEYS.DB_ID)          || '';
  get('status-property').value = localStorage.getItem(STORAGE_KEYS.STATUS_PROP)    || 'Status';
  get('default-status').value  = localStorage.getItem(STORAGE_KEYS.DEFAULT_STATUS) || 'To Do';
}

function saveSettings() {
  const token   = get('notion-token').value.trim();
  const dbId    = get('notion-db-id').value.trim().replace(/-/g, '');
  const statusP = get('status-property').value.trim() || 'Status';
  const defStat = get('default-status').value.trim()  || 'To Do';

  if (!token || !dbId) {
    showSettingsStatus('Please enter both the Integration Token and Database ID.', 'error');
    return;
  }
  localStorage.setItem(STORAGE_KEYS.TOKEN,          token);
  localStorage.setItem(STORAGE_KEYS.DB_ID,          dbId);
  localStorage.setItem(STORAGE_KEYS.STATUS_PROP,    statusP);
  localStorage.setItem(STORAGE_KEYS.DEFAULT_STATUS, defStat);

  showSettingsStatus('Settings saved!', 'success');
  setTimeout(() => {
    hide('settings-view');
    show('main-view');
    hide('not-configured');
    loadEmailData().then(() => { show('email-preview'); show('task-form'); });
  }, 900);
}

function showSettingsStatus(msg, type) {
  const el = get('settings-status');
  el.textContent  = msg;
  el.className    = 'status-msg ' + type;
  show('settings-status');
}

// ── Load email data ───────────────────────────────────────────
async function loadEmailData() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    const data = {
      subject:   item.subject || '(No Subject)',
      from:      '',
      fromEmail: '',
      body:      '',
      itemId:    item.itemId  || '',
      webLink:   item.webLink || ''
    };
    if (item.from) {
      data.from      = item.from.displayName  || '';
      data.fromEmail = item.from.emailAddress || '';
    }
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: data }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        result.asyncContext.body = result.value.substring(0, 1500).trim();
      }
      currentEmail = result.asyncContext;
      get('preview-from').textContent    = currentEmail.from
        ? 'From: ' + currentEmail.from + ' <' + currentEmail.fromEmail + '>'
        : 'Unknown sender';
      get('preview-subject').textContent = currentEmail.subject;
      get('task-title').value            = currentEmail.subject;
      resolve(currentEmail);
    });
  });
}

function buildEmailLink(email) {
  if (email.webLink && email.webLink.startsWith('http')) return email.webLink;
  return 'https://outlook.office.com/mail/deeplink?ItemID=' + encodeURIComponent(email.itemId);
}

// ── Create Notion Task ────────────────────────────────────────
async function createNotionTask(quickAdd) {
  if (!currentEmail) { showError('Could not read email. Please reopen the panel.'); return; }

  const token      = localStorage.getItem(STORAGE_KEYS.TOKEN);
  const dbId       = localStorage.getItem(STORAGE_KEYS.DB_ID);
  const statusProp = localStorage.getItem(STORAGE_KEYS.STATUS_PROP)    || 'Status';
  const defStatus  = localStorage.getItem(STORAGE_KEYS.DEFAULT_STATUS) || 'To Do';

  const title       = quickAdd ? currentEmail.subject : (get('task-title').value.trim() || currentEmail.subject);
  const description = quickAdd ? '' : get('task-description').value.trim();
  const dueDate     = quickAdd ? '' : get('task-due-date').value;
  const priority    = quickAdd ? '' : get('task-priority').value;
  const inclLink    = quickAdd ? true : get('include-email-link').checked;
  const inclBody    = quickAdd ? false : get('include-body-preview').checked;

  if (!token || !dbId) { showError('Please configure settings first.'); show('settings-view'); hide('main-view'); return; }

  setLoading(true);
  hide('error-msg');

  try {
    const emailLink = buildEmailLink(currentEmail);
    const pageUrl   = await postToNotion(token, dbId, {
      title, description, dueDate, priority,
      inclLink, inclBody, emailLink, statusProp, defStatus
    });
    setLoading(false);
    hide('task-form');
    hide('email-preview');
    get('notion-link').href = pageUrl;
    show('success-msg');
  } catch (err) {
    setLoading(false);
    showError(err.message || 'Failed to create task. Check your settings.');
  }
}

// ── Notion API (direct fetch via CORS proxy) ─────────────────
async function postToNotion(token, dbId, opts) {
  const { title, description, dueDate, priority,
          inclLink, inclBody, emailLink, statusProp, defStatus } = opts;

  const properties = {
    'Name': { title: [{ text: { content: title } }] },
    [statusProp]: { select: { name: defStatus } }
  };
  if (dueDate)  properties['Due Date'] = { date: { start: dueDate } };
  if (priority) properties['Priority'] = { select: { name: priority } };

  const children = [];

  if (inclLink) {
    children.push({
      object: 'block', type: 'callout',
      callout: {
        rich_text: [
          { type: 'text', text: { content: 'Source Email: ' } },
          { type: 'text',
            text: { content: currentEmail.subject, link: { url: emailLink } },
            annotations: { bold: true, color: 'blue' }
          }
        ],
        icon: { emoji: '📧' },
        color: 'blue_background'
      }
    });
    children.push({
      object: 'block', type: 'paragraph',
      paragraph: {
        rich_text: [{
          type: 'text',
          text: { content: 'From: ' + currentEmail.from + ' <' + currentEmail.fromEmail + '>' },
          annotations: { color: 'gray' }
        }]
      }
    });
    children.push({ object: 'block', type: 'divider', divider: {} });
  }

  if (description) {
    children.push({
      object: 'block', type: 'heading_3',
      heading_3: { rich_text: [{ type: 'text', text: { content: 'Notes' } }] }
    });
    children.push({
      object: 'block', type: 'paragraph',
      paragraph: { rich_text: [{ type: 'text', text: { content: description } }] }
    });
  }

  if (inclBody && currentEmail.body) {
    children.push({
      object: 'block', type: 'heading_3',
      heading_3: { rich_text: [{ type: 'text', text: { content: 'Email Preview' } }] }
    });
    children.push({
      object: 'block', type: 'quote',
      quote: { rich_text: [{ type: 'text', text: { content: currentEmail.body.substring(0, 1500) } }] }
    });
  }

  const payload = {
    parent: { database_id: dbId },
    properties,
    children
  };

  // Use corsproxy.io to bypass CORS — the token never leaves your browser
  const targetUrl = CORS_PROXY + encodeURIComponent(NOTION_API);

  const res = await fetch(targetUrl, {
    method: 'POST',
    headers: {
      'Content-Type':    'application/json',
      'Authorization':   'Bearer ' + token,
      'Notion-Version':  '2022-06-28'
    },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.message || 'Notion API error ' + res.status);
  }

  const data = await res.json();
  return data.url;
}