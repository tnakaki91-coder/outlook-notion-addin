/* ============================================================
   Local HTTPS Proxy Server for Notion API
   server.js  --  run with: node server.js
   ============================================================ */

const https = require('https');
const http  = require('http');
const fs    = require('fs');
const path  = require('path');

const PORT = 3000;

const SSL_KEY  = path.join(__dirname, 'certs', 'server.key');
const SSL_CERT = path.join(__dirname, 'certs', 'server.cert');

const MIME_TYPES = {
  '.html': 'text/html',
  '.css':  'text/css',
  '.js':   'application/javascript',
  '.png':  'image/png',
  '.json': 'application/json',
  '.xml':  'application/xml'
};

function setCORS(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

async function handler(req, res) {
  setCORS(res);

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/notion-proxy') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { token, payload } = JSON.parse(body);
        if (!token || !payload) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ message: 'Missing token or payload' }));
          return;
        }
        const result = await forwardToNotion(token, payload);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: err.message }));
      }
    });
    return;
  }

  let filePath = req.url === '/' ? '/taskpane.html' : req.url;
  filePath = path.join(__dirname, filePath);

  if (!fs.existsSync(filePath)) {
    res.writeHead(404);
    res.end('Not found');
    return;
  }

  const ext  = path.extname(filePath);
  const mime = MIME_TYPES[ext] || 'application/octet-stream';
  res.writeHead(200, { 'Content-Type': mime });
  fs.createReadStream(filePath).pipe(res);
}

function forwardToNotion(token, payload) {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify(payload);
    const options = {
      hostname: 'api.notion.com',
      port: 443,
      path: '/v1/pages',
      method: 'POST',
      headers: {
        'Authorization':  'Bearer ' + token,
        'Content-Type':   'application/json',
        'Notion-Version': '2022-06-28',
        'Content-Length': Buffer.byteLength(postData)
      }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode >= 400) {
            reject(new Error(parsed.message || 'Notion error ' + res.statusCode));
          } else {
            resolve(parsed);
          }
        } catch (e) {
          reject(new Error('Invalid JSON from Notion API'));
        }
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

try {
  const sslOptions = {
    key:  fs.readFileSync(SSL_KEY),
    cert: fs.readFileSync(SSL_CERT)
  };
  https.createServer(sslOptions, handler).listen(PORT, () => {
    console.log('HTTPS server running at https://localhost:' + PORT);
  });
} catch (e) {
  console.error('SSL certs not found. Run: npm run gen-cert');
  console.log('Starting HTTP server for testing only (will NOT work with Outlook)...');
  http.createServer(handler).listen(PORT, () => {
    console.log('HTTP server running at http://localhost:' + PORT);
  });
}