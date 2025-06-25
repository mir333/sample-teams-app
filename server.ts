import express, { Request, Response } from 'express';
import axios from 'axios';
import https from 'https';
import fs from 'fs';
import path from 'path';
import cors from 'cors';

const app = express();
app.use(cors());
app.use(express.json());

// Replace with your Azure AD app registration values:
const CLIENT_ID = 'YOUR_CLIENT_ID';
const CLIENT_SECRET = 'YOUR_CLIENT_SECRET';
const TENANT_ID = 'YOUR_TENANT_ID';

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

app.post('/api/exchange-token', async (req: Request, res: Response) => {
  const { token } = req.body;
  if (!token) {
    return res.status(400).json({ error: 'Missing token' });
  }

  try {
    // OBO token exchange
    const params = new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      requested_token_use: 'on_behalf_of',
      scope: GRAPH_SCOPE,
      assertion: token,
    });

    const response = await axios.post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      params,
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    return res.json({ graphToken: response.data.access_token });
  } catch (e: any) {
    console.error(e.response?.data || e.message);
    return res.status(500).json({ error: 'Token exchange failed', details: e.response?.data || e.message });
  }
});

const PORT = 3001;

// HTTPS server setup
const certDir = path.resolve(__dirname, 'certs');
const key = fs.readFileSync(path.join(certDir, 'localhost-key.pem'));
const cert = fs.readFileSync(path.join(certDir, 'localhost.pem'));

https.createServer({ key, cert }, app).listen(PORT, () => {
  console.log(`HTTPS backend listening on port ${PORT}`);
});
