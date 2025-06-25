import { useEffect, useState } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import reactLogo from './assets/react.svg';
import viteLogo from '/vite.svg';
import './App.css';

function App() {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [user, setUser] = useState<any>(null);
  const [graphData, setGraphData] = useState<any>(null);
  const [count, setCount] = useState(0);

  useEffect(() => {
    microsoftTeams.app.initialize()
      .then(async () => {
        try {
          // Use Promise-based getAuthToken
          const token = await microsoftTeams.authentication.getAuthToken({
            resources: ["https://graph.microsoft.com"]
          });
          // Send token to backend to exchange for Graph access token
          const res = await fetch('http://localhost:3001/api/exchange-token', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token }),
          });
          const data = await res.json();
          if (data.graphToken) {
            // Call Graph API
            const meRes = await fetch('https://graph.microsoft.com/v1.0/me', {
              headers: { Authorization: `Bearer ${data.graphToken}` },
            });
            const me = await meRes.json();
            setUser(me);
            setGraphData(me);
            setLoading(false);
          } else {
            setError('No Graph token returned from backend.');
            setLoading(false);
          }
        } catch (e: any) {
          setError('Token exchange or Graph call failed: ' + e.message);
          setLoading(false);
        }
      })
      .catch((e) => {
        setError('Teams SDK initialization failed: ' + e.message);
        setLoading(false);
      });
  }, []);

  if (loading) return <div>Loading Teams SSO...</div>;
  if (error) return <div style={{color: 'red'}}>Error: {error}</div>;
  return (
    <div style={{ padding: 40, fontFamily: 'sans-serif', textAlign: 'center' }}>
      <h2>Welcome, {user?.displayName}</h2>
      <div style={{ marginTop: 20 }}>
        <h3>Microsoft Graph Data:</h3>
        <pre style={{ textAlign: 'left', marginTop: 20 }}>{JSON.stringify(graphData, null, 2)}</pre>
      </div>
      <div>
        <a href="https://vite.dev" target="_blank">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>Vite + React</h1>
      <div className="card">
        <button onClick={() => setCount((count) => count + 1)}>
          count is {count}
        </button>
        <p>
          Edit <code>src/App.tsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">
        Click on the Vite and React logos to learn more
      </p>
    </div>
  );
}

export default App;
