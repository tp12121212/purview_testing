import { useEffect, useMemo, useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig.js';

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL ?? 'http://localhost:4000';
const authMode = (import.meta.env.VITE_AUTH_MODE ?? 'msal').toLowerCase();
const isMsalAuth = authMode === 'msal';

const formatJson = (value) => JSON.stringify(value, null, 2);

const ResultCard = ({ title, content }) => {
  if (!content) {
    return null;
  }

  return (
    <section className="card">
      <h3>{title}</h3>
      <pre>{formatJson(content)}</pre>
    </section>
  );
};

const StreamTable = ({ streams }) => {
  if (!streams?.length) {
    return null;
  }

  return (
    <section className="card">
      <h3>Extracted Streams</h3>
      <div className="table-wrapper">
        <table>
          <thead>
            <tr>
              <th>Index</th>
              <th>Kind</th>
              <th>Name</th>
              <th>Source File</th>
            </tr>
          </thead>
          <tbody>
            {streams.map((stream) => (
              <tr key={`${stream.StreamIndex}-${stream.Name}`}>
                <td>{stream.StreamIndex}</td>
                <td>{stream.Kind}</td>
                <td>{stream.Name}</td>
                <td>{stream.SourceFile}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
};

const DataClassificationCards = ({ data }) => {
  if (!data?.length) {
    return null;
  }

  return (
    <section className="card">
      <h3>Classification Results</h3>
      <div className="results-grid">
        {data.map((item) => (
          <article key={`${item.StreamIndex}-${item.Name}`}>
            <header>
              <strong>{item.Name}</strong>
              <span>{item.Kind}</span>
            </header>
            <pre>{formatJson(item.Result)}</pre>
          </article>
        ))}
      </div>
    </section>
  );
};

const useAccessToken = (enabled) => {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const getToken = async () => {
    if (!enabled) {
      return null;
    }
    if (!account) {
      throw new Error('No signed in account.');
    }
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account
    });
    return response.accessToken;
  };

  return { account, getToken };
};

const buildFormData = (file, payload) => {
  const form = new FormData();
  form.append('file', file);
  Object.entries(payload).forEach(([key, value]) => {
    if (typeof value !== 'undefined') {
      form.append(key, value);
    }
  });
  return form;
};

export default function App() {
  const { instance } = useMsal();
  const { account, getToken } = useAccessToken(isMsalAuth);
  const [file, setFile] = useState(null);
  const [result, setResult] = useState(null);
  const [classification, setClassification] = useState(null);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [sits, setSits] = useState([]);
  const [selectedSits, setSelectedSits] = useState('');
  const [useAllSits, setUseAllSits] = useState(true);
  const [userPrincipalName, setUserPrincipalName] = useState('');

  const isAuthenticated = isMsalAuth ? Boolean(account) : Boolean(userPrincipalName);

  const selectionHint = useMemo(() => {
    if (!sits.length) {
      return 'Load Sensitive Information Types after signing in.';
    }
    return 'Enter comma-separated SIT display names or IDs, or leave blank to use all.';
  }, [sits.length]);

  const handleLogin = async () => {
    await instance.loginPopup(loginRequest);
  };

  const handleLogout = async () => {
    await instance.logoutPopup({
      account
    });
  };

  const fetchSits = async () => {
    setError('');
    try {
      const token = await getToken();
      const response = await fetch(`${apiBaseUrl}/api/sensitive-information-types`, {
        headers: {
          ...(isMsalAuth ? { Authorization: `Bearer ${token}` } : { 'X-User-Principal-Name': userPrincipalName })
        }
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error ?? 'Failed to load SITs.');
      }
      setSits(payload.items ?? []);
    } catch (err) {
      setError(err.message);
    }
  };

  useEffect(() => {
    if (!isAuthenticated) {
      setSits([]);
    }
  }, [isAuthenticated]);

  const handleRequest = async (endpoint, setState) => {
    if (!file) {
      setError('Please select a file to upload.');
      return;
    }
    if (!isMsalAuth && !userPrincipalName) {
      setError('Please enter your user principal name.');
      return;
    }
    setLoading(true);
    setError('');
    setState(null);

    try {
      const token = await getToken();
      const payload = buildFormData(file, {
        selectedSits: useAllSits ? '' : selectedSits,
        useAllSits,
        userPrincipalName: isMsalAuth ? undefined : userPrincipalName
      });

      const response = await fetch(`${apiBaseUrl}${endpoint}`, {
        method: 'POST',
        headers: {
          ...(isMsalAuth ? { Authorization: `Bearer ${token}` } : { 'X-User-Principal-Name': userPrincipalName })
        },
        body: payload
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error ?? 'Request failed.');
      }

      setState(data.result ?? null);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app">
      <header className="hero">
        <div>
          <p className="eyebrow">Microsoft Purview</p>
          <h1>Extraction & Classification</h1>
          <p>
            Run Test-TextExtraction and Test-DataClassification with secure Microsoft 365 sign-in.
          </p>
        </div>
        <div className="auth-actions">
          {isMsalAuth && isAuthenticated ? (
            <>
              <div className="user-pill">
                <span>{account?.name ?? account?.username}</span>
              </div>
              <button type="button" className="secondary" onClick={handleLogout}>
                Sign out
              </button>
            </>
          ) : isMsalAuth ? (
            <button type="button" onClick={handleLogin}>
              Sign in with Microsoft 365
            </button>
          ) : (
            <div className="user-pill">
              <span>Device code / user auth</span>
            </div>
          )}
        </div>
      </header>

      <main>
        <section className="card">
          <h2>Upload & Run</h2>
          <div className="field">
            <label htmlFor="file">Document or Email File</label>
            <input
              id="file"
              type="file"
              accept=".pdf,.msg,.eml,.docx"
              onChange={(event) => setFile(event.target.files?.[0] ?? null)}
            />
          </div>

          {!isMsalAuth && (
            <div className="field">
              <label htmlFor="upn">User Principal Name</label>
              <input
                id="upn"
                type="text"
                placeholder="user@contoso.com"
                value={userPrincipalName}
                onChange={(event) => setUserPrincipalName(event.target.value)}
              />
              <p className="status">
                You will be prompted for device code authentication on the server console.
              </p>
            </div>
          )}

          <div className="field">
            <label>Sensitive Information Types</label>
            <div className="sits-actions">
              <button type="button" className="secondary" onClick={fetchSits} disabled={!isAuthenticated}>
                Load SIT Catalog
              </button>
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={useAllSits}
                  onChange={(event) => setUseAllSits(event.target.checked)}
                />
                Use all SITs
              </label>
            </div>
            <input
              type="text"
              placeholder={selectionHint}
              value={selectedSits}
              disabled={useAllSits}
              onChange={(event) => setSelectedSits(event.target.value)}
            />
            {sits.length > 0 && (
              <div className="sits-list">
                {sits.slice(0, 10).map((sit) => (
                  <span key={sit.Id}>{sit.Display}</span>
                ))}
                {sits.length > 10 && <span>+{sits.length - 10} more</span>}
              </div>
            )}
          </div>

          <div className="actions">
            <button type="button" onClick={() => handleRequest('/api/extraction', setResult)} disabled={!isAuthenticated || loading}>
              Run Extraction
            </button>
            <button type="button" className="secondary" onClick={() => handleRequest('/api/classification', setClassification)} disabled={!isAuthenticated || loading}>
              Run Classification
            </button>
          </div>
          {loading && <p className="status">Working on your request...</p>}
          {error && <p className="error">{error}</p>}
        </section>

        <StreamTable streams={result?.Streams ?? classification?.Streams} />
        <ResultCard title="Extraction Details" content={result?.Extraction} />
        <DataClassificationCards data={classification?.DataClassification} />
        <ResultCard title="Classification Details" content={classification?.DataClassification} />
      </main>

      <footer>
        <p>
          Ensure the backend is running with the required PowerShell modules installed for Exchange Online and Purview compliance.
        </p>
      </footer>
    </div>
  );
}
