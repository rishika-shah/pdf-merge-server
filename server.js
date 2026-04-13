const express = require('express');
const cors = require('cors');
const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const { PDFDocument } = require('pdf-lib');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ─────────────────────────────────────────
// MICROSOFT GRAPH CONFIG
// ─────────────────────────────────────────
const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

let accessToken = null;
let tokenExpiry = 0;

// Get (or reuse) a Graph API access token
async function getAccessToken() {
  if (accessToken && Date.now() < tokenExpiry - 60000) {
    return accessToken;
  }

  console.log('Fetching new access token...');
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });

  const res = await fetch(url, { method: 'POST', body });
  const json = await res.json();

  if (!res.ok || !json.access_token) {
    throw new Error(`Failed to get access token: ${json.error_description || json.error}`);
  }

  accessToken = json.access_token;
  tokenExpiry = Date.now() + (json.expires_in * 1000);
  console.log('Access token obtained, expires in', json.expires_in, 'seconds');
  return accessToken;
}

// ─────────────────────────────────────────
// FETCH PDF VIA GRAPH API
// ─────────────────────────────────────────
async function fetchPdfViaGraph(sharingUrl) {
  const token = await getAccessToken();

  // Convert OneDrive sharing URL to Graph API encoded sharing URL
  // See: https://learn.microsoft.com/en-us/graph/api/shares-get
  const base64 = Buffer.from(sharingUrl).toString('base64')
    .replace(/=/g, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
  const encodedUrl = `u!${base64}`;

  // First get the file metadata to get the download URL
  const metaUrl = `https://graph.microsoft.com/v1.0/shares/${encodedUrl}/driveItem`;
  console.log(`Getting file metadata for: ${sharingUrl.substring(0, 60)}...`);

  const metaRes = await fetch(metaUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!metaRes.ok) {
    const err = await metaRes.json().catch(() => ({}));
    throw new Error(`Graph metadata error ${metaRes.status}: ${err.error?.message || JSON.stringify(err)}`);
  }

  const meta = await metaRes.json();
  const downloadUrl = meta['@microsoft.graph.downloadUrl'];

  if (!downloadUrl) {
    throw new Error(`No download URL found for: ${sharingUrl}`);
  }

  console.log(`Downloading: ${meta.name || 'unknown'} (${meta.size || '?'} bytes)`);

  // Download the actual file
  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) {
    throw new Error(`Download failed: HTTP ${fileRes.status}`);
  }

  const buffer = await fileRes.arrayBuffer();
  return new Uint8Array(buffer);
}

// ─────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────
app.get('/', (req, res) => res.send('PDF Merge Server is running.'));

app.post('/merge', async (req, res) => {
  const { urls, filename } = req.body;

  if (!urls || !Array.isArray(urls) || urls.length === 0) {
    return res.status(400).json({ error: 'No URLs provided' });
  }

  console.log(`\nMerging ${urls.length} PDFs...`);

  const mergedPdf = await PDFDocument.create();
  const errors = [];

  for (const url of urls) {
    // Strip the &download=1 we added in the dashboard — Graph API doesn't need it
    const cleanUrl = url.replace(/[&?]download=1/, '');

    try {
      const pdfBytes = await fetchPdfViaGraph(cleanUrl);

      let srcDoc;
      try {
        srcDoc = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
      } catch (e) {
        throw new Error(`Could not parse PDF: ${e.message}`);
      }

      const pages = await mergedPdf.copyPagesFrom(srcDoc, srcDoc.getPageIndices());
      pages.forEach(page => mergedPdf.addPage(page));
      console.log(`✓ Added ${pages.length} page(s)`);

    } catch (e) {
      console.warn(`✗ Failed for ${cleanUrl.substring(0, 60)}...: ${e.message}`);
      errors.push({ url: cleanUrl, error: e.message });
    }
  }

  const pageCount = mergedPdf.getPageCount();

  if (pageCount === 0) {
    return res.status(422).json({
      error: 'No pages could be merged.',
      details: errors,
    });
  }

  const mergedBytes = await mergedPdf.save();
  const outputFilename = filename || 'styles-export.pdf';

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
  res.setHeader('Content-Length', mergedBytes.length);
  res.send(Buffer.from(mergedBytes));

  console.log(`✓ Done. Merged PDF: ${pageCount} pages. ${errors.length} failed.\n`);
});

// ─────────────────────────────────────────
// START
// ─────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
