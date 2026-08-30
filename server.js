import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static(process.cwd()));
app.use(express.static(__dirname));

// Explicit asset routes to ensure Vercel serverless bundles can resolve them
app.get('/icon.png', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'icon.png'));
});

app.get('/kop-surat.png', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'kop-surat.png'));
});

app.get('/manifest.json', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'manifest.json'));
});

app.get('/sw.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript');
  res.sendFile(path.join(process.cwd(), 'sw.js'));
});

// Explicit route fallbacks
app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'dashboard_admin.html'));
});

app.get('/payroll', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'payroll.html'));
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Fallback to index.html for root or unknown
app.get('/', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running at http://0.0.0.0:${PORT}`);
});
