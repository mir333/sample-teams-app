import React, { useState } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider, useMsal, useIsAuthenticated } from '@azure/msal-react';
import reactLogo from './assets/react.svg'
import viteLogo from '/vite.svg'
import './App.css'

const msalConfig = {
  auth: {
    clientId: 'e0bf234d-7550-4fa3-bc00-2d4c21276f1c',
    authority: 'https://login.microsoftonline.com/d3d3b20f-00ce-4a0f-9975-d79117daa055',
    redirectUri: window.location.origin,
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

function LoginContent() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [graphData, setGraphData] = useState<any>(null);
  const [count, setCount] = useState(0);

  const login = async () => {
    try {
      const loginResponse = await instance.loginPopup({
        scopes: ["User.Read"]
      });
      if (loginResponse.account) {
        instance.setActiveAccount(loginResponse.account);
      }
    } catch (e) {
      alert('Login failed: ' + (e as any).message);
    }
  };

  const logout = () => {
    instance.logoutPopup();
  };

  const callGraph = async () => {
    const account = instance.getActiveAccount();
    if (!account) return;
    try {
      const result = await instance.acquireTokenSilent({
        account,
        scopes: ["User.Read"]
      });
      const res = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${result.accessToken}` }
      });
      const data = await res.json();
      setGraphData(data);
    } catch (e) {
      alert('Failed to call Graph: ' + (e as any).message);
    }
  };

  if (!isAuthenticated) {
    return (
      <>
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
        <button onClick={login}>Login with Microsoft</button>
      </>
    );
  }

  return (
    <div style={{ padding: 40, fontFamily: 'sans-serif', textAlign: 'center' }}>
      <h2>Welcome, {accounts[0]?.username}</h2>
      <button onClick={logout}>Logout</button>
      <div style={{ marginTop: 20 }}>
        <button onClick={callGraph}>Call Microsoft Graph</button>
        {graphData && (
          <pre style={{ textAlign: 'left', marginTop: 20 }}>{JSON.stringify(graphData, null, 2)}</pre>
        )}
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

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <LoginContent />
    </MsalProvider>
  );
}
