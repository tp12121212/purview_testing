import { useEffect, useMemo, useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { adminConsentScopes, authorityHost, clientId, complianceScope, exoScope, loginRequest } from './authConfig.js';

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL ?? 'http://localhost:4000';

const formatJson = (value) => JSON.stringify(value, null, 2);

const normalizeErrorMessage = (error) => {
  if (!error) {
    return '';
  }
  if (typeof error === 'string') {
    return error;
  }
  return error.errorMessage ?? error.message ?? JSON.stringify(error);
};

const extractResourceFromError = (message) => {
  if (!message) {
    return null;
  }
  const match = message.match(/Resource value from request:\s*(\S+)/i);
  return match?.[1] ?? null;
};

const buildAdminConsentUrl = (tenantId) => {
  if (!clientId) {
    return null;
  }
  const tenantSegment = tenantId || 'organizations';
  const redirectUri = encodeURIComponent(window.location.origin);
  if (!adminConsentScopes.length) {
    return null;
  }
  const scopeParam = encodeURIComponent(adminConsentScopes.join(' '));
  return `${authorityHost}/${tenantSegment}/v2.0/adminconsent?client_id=${clientId}&scope=${scopeParam}&redirect_uri=${redirectUri}`;
};

const describeAuthError = (error, tenantId) => {
  const message = normalizeErrorMessage(error);
  if (!message) {
    return null;
  }

  const lowered = message.toLowerCase();
  const resource = extractResourceFromError(message);
  const adminConsentUrl = buildAdminConsentUrl(tenantId);

  if (lowered.includes('aadsts650057') || lowered.includes('invalid_resource')) {
    return {
      title: 'Compliance resource not available for this tenant',
      summary: 'This tenant cannot issue a compliance token for the app. Ask a tenant admin to grant consent, and confirm Purview/Compliance is provisioned.',
      description: 'The tenant either has not granted admin consent for this app or does not expose Purview/Compliance delegated scopes yet.',
      resource,
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts65001') || lowered.includes('need admin approval')) {
    return {
      title: 'Admin consent required',
      summary: 'A tenant admin must approve this app before compliance or Exchange tokens can be issued.',
      description: 'Ask a tenant admin to grant consent for Exchange Online and Microsoft Purview compliance delegated permissions.',
      resource,
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts70011') && lowered.includes('scope')) {
    return {
      title: 'Invalid admin consent scopes',
      summary: 'The admin consent request included invalid scopes. Only resource scopes (/.default) are allowed.',
      description: 'Ask the app owner to ensure the admin consent link includes only resource /.default scopes.',
      resource,
      adminConsentUrl,
      details: message
    };
  }

  return null;
};

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

const useAccessToken = () => {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const getTokenForScope = async (scope) => {
    if (!account) {
      throw new Error('No signed in account.');
    }
    try {
      const response = await instance.acquireTokenSilent({
        scopes: [scope],
        account
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await instance.acquireTokenPopup({
          scopes: [scope]
        });
        return response.accessToken;
      }
      throw error;
    }
  };

  return { account, getTokenForScope };
};

const buildFormData = (file, payload) => {
  const form = new FormData();
  form.append('file', file);
  Object.entries(payload).forEach(([key, value]) => {
    if (typeof value === 'undefined' || value === null) {
      return;
    }
    if (typeof value === 'string' && value.trim() === '') {
      return;
    }
    form.append(key, value);
  });
  return form;
};

export default function App() {
  const { instance } = useMsal();
  const { account, getTokenForScope } = useAccessToken();
  const [file, setFile] = useState(null);
  const [result, setResult] = useState(null);
  const [classification, setClassification] = useState(null);
  const [error, setError] = useState('');
  const [authHelp, setAuthHelp] = useState(null);
  const [loading, setLoading] = useState(false);
  const [consentLoading, setConsentLoading] = useState(false);
  const [sits, setSits] = useState([]);
  const [selectedSits, setSelectedSits] = useState('');
  const [useAllSits, setUseAllSits] = useState(true);

  const isAuthenticated = Boolean(account);
  const tenantId = account?.tenantId ?? null;

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

  const handleGrantPermissions = async () => {
    setAuthHelp(null);
    setError('');
    setConsentLoading(true);
    try {
      await getTokenForScope(exoScope);
      await getTokenForScope(complianceScope);
    } catch (err) {
      const help = describeAuthError(err, tenantId);
      setAuthHelp(help);
      setError(help?.summary ?? err.message);
    } finally {
      setConsentLoading(false);
    }
  };

  const fetchSits = async () => {
    setError('');
    setAuthHelp(null);
    try {
      const token = await getTokenForScope(complianceScope);
      const response = await fetch(`${apiBaseUrl}/api/sensitive-information-types`, {
        headers: {
          Authorization: `Bearer ${token}`
        }
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error ?? 'Failed to load SITs.');
      }
      setSits(payload.items ?? []);
    } catch (err) {
      const help = describeAuthError(err, tenantId);
      setAuthHelp(help);
      setError(help?.summary ?? err.message);
    }
  };

  useEffect(() => {
    if (!isAuthenticated) {
      setSits([]);
    }
    if (authHelp) {
      setAuthHelp(null);
    }
  }, [isAuthenticated]);

  const handleRequest = async (endpoint, setState, options = {}) => {
    if (!file) {
      setError('Please select a file to upload.');
      return;
    }
    setLoading(true);
    setError('');
    setAuthHelp(null);
    setState(null);

    try {
      const needsExchange = Boolean(options.needsExchangeToken);
      const needsCompliance = Boolean(options.needsComplianceToken);
      const exchangeToken = needsExchange ? await getTokenForScope(exoScope) : null;
      const complianceToken = needsCompliance ? await getTokenForScope(complianceScope) : null;
      const payload = buildFormData(file, {
        selectedSits: useAllSits ? '' : selectedSits,
        useAllSits
      });

      const authHeaderToken = complianceToken ?? exchangeToken;
      const response = await fetch(`${apiBaseUrl}${endpoint}`, {
        method: 'POST',
        headers: {
          ...(authHeaderToken ? { Authorization: `Bearer ${authHeaderToken}` } : {}),
          ...(exchangeToken && exchangeToken !== authHeaderToken ? { 'X-Exchange-Token': exchangeToken } : {})
        },
        body: payload
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error ?? 'Request failed.');
      }

      setState(data.result ?? null);
    } catch (err) {
      const help = describeAuthError(err, tenantId);
      setAuthHelp(help);
      setError(help?.summary ?? err.message);
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
          {isAuthenticated ? (
            <>
              <div className="user-pill">
                <span>{account?.name ?? account?.username}</span>
              </div>
              <button type="button" className="secondary" onClick={handleLogout}>
                Sign out
              </button>
            </>
          ) : (
            <button type="button" onClick={handleLogin}>
              Sign in with Microsoft 365
            </button>
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
            <button
              type="button"
              onClick={() => handleRequest('/api/extraction', setResult, { needsExchangeToken: true })}
              disabled={!isAuthenticated || loading}
            >
              Run Extraction
            </button>
            <button
              type="button"
              className="secondary"
              onClick={() => handleRequest('/api/classification', setClassification, { needsExchangeToken: true, needsComplianceToken: true })}
              disabled={!isAuthenticated || loading}
            >
              Run Classification
            </button>
          </div>
          {loading && <p className="status">Working on your request...</p>}
          {consentLoading && <p className="status">Requesting permissions...</p>}
          {error && <p className="error">{error}</p>}
        </section>

        {isAuthenticated && (
          <section className="card notice">
            <h3>Tenant Setup</h3>
            <p>
              First-time users should grant Exchange and Compliance permissions in the browser. Tenant admins can also grant
              consent on behalf of the organization.
            </p>
            <div className="actions">
              <button type="button" onClick={handleGrantPermissions} disabled={consentLoading}>
                Grant permissions
              </button>
              {buildAdminConsentUrl(tenantId) && (
                <button
                  type="button"
                  className="secondary"
                  onClick={() => window.open(buildAdminConsentUrl(tenantId), '_blank', 'noopener,noreferrer')}
                >
                  Admin consent for organization
                </button>
              )}
            </div>
          </section>
        )}

        {authHelp && (
          <section className="card notice">
            <h3>{authHelp.title}</h3>
            <p>{authHelp.description}</p>
            {authHelp.resource && (
              <p>
                <strong>Requested resource:</strong> {authHelp.resource}
              </p>
            )}
            {authHelp.adminConsentUrl && (
              <button
                type="button"
                className="secondary"
                onClick={() => window.open(authHelp.adminConsentUrl, '_blank', 'noopener,noreferrer')}
              >
                Open admin consent link
              </button>
            )}
            {authHelp.details && <pre>{authHelp.details}</pre>}
          </section>
        )}

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
