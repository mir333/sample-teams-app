import express, { Request, Response } from 'express';
import axios from 'axios';

const app = express();
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
app.listen(PORT, () => console.log(`Backend listening on port ${PORT}`));
