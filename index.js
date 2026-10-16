const express = require('express');
const path = require('path');

const app = express();

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// 1. Mount API Handlers
app.all('/api/chat-ai', (req, res) => require('./api/chat-ai')(req, res));
app.all('/api/sync-keep', (req, res) => require('./api/sync-keep')(req, res));
app.all('/api/monthly-archives', (req, res) => require('./api/monthly-archives')(req, res));
app.all('/api/monthly-targets', (req, res) => require('./api/monthly-targets')(req, res));
app.all('/api/sales-photos', (req, res) => require('./api/sales-photos')(req, res));
app.all('/api/save-excel', (req, res) => require('./api/save-excel')(req, res));

// 2. Serve all static assets (CSS, JS, images, fonts, XLSX, CSV)
app.use(express.static(__dirname));

// 3. SPA fallback for root and page routes (only non-api routes!)
app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({ error: 'API route not found' });
  }
  res.sendFile(path.join(__dirname, 'index.html'));
});

module.exports = app;
