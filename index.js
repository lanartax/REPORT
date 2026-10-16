const express = require('express');
const path = require('path');

const app = express();

// Serve all static assets (CSS, JS, images, fonts, XLSX, CSV)
app.use(express.static(__dirname));

// SPA fallback for root and page routes
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

module.exports = app;
