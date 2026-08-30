import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static(__dirname));

// Explicit route fallbacks
app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'dashboard_admin.html'));
});

app.get('/payroll', (req, res) => {
  res.sendFile(path.join(__dirname, 'payroll.html'));
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Fallback to index.html for root or unknown
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running at http://0.0.0.0:${PORT}`);
});
