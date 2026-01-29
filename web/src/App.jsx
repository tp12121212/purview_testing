import { useEffect, useMemo, useState } from 'react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useMsal } from '@azure/msal-react';
import {
  buildGraphAdminConsentUrl,
  buildLoginRequest,
  buildScopeList,
  formatScopeList,
  summarizeScopeSet
} from './authConfig.js';
import { createGraphClient } from './graphClient.js';
import {
  extractTextFromBuffer,
  extractTextFromFile,
  extractTextFromHtml
} from './extractors.js';
import {
  aggregateClassificationResults,
  buildGraphClassificationResults,
  detectSensitiveInfo,
  evaluateSitDetectors,
  findInvalidDetectors,
  getSampleDetectors
} from './classification.js';
import { parseRulePackXml } from './rulePackParser.js';

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

const extractErrorCode = (error) => error?.errorCode ?? error?.code ?? '';

const shouldUseRedirect = (message, code) => {
  const lowered = message.toLowerCase();
  return code?.includes('popup') || lowered.includes('popup') || lowered.includes('blocked');
};

const isPopupOrCookieIssue = (message, code) => {
  const lowered = message.toLowerCase();
  return (
    code?.includes('popup')
    || lowered.includes('popup')
    || lowered.includes('third-party')
    || lowered.includes('third party')
    || lowered.includes('cookie')
  );
};

const describeAuthError = (error, tenantId, runtimeConfig) => {
  const message = normalizeErrorMessage(error);
  if (!message) {
    return null;
  }

  const lowered = message.toLowerCase();
  const adminConsentUrl = buildGraphAdminConsentUrl(runtimeConfig, tenantId);

  if (lowered.includes('aadsts70011') && lowered.includes('scope')) {
    return {
      title: 'Invalid scopes requested',
      summary: 'The requested scopes are not valid. Login scopes must only include openid/profile/email.',
      description: 'Check your Login Scopes setting and ensure Graph scopes are requested separately.',
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts65001') || lowered.includes('need admin approval') || lowered.includes('consent_required')) {
    return {
      title: 'Admin consent required',
      summary: 'A tenant admin must approve this app before some Graph permissions can be issued.',
      description: 'Ask a tenant admin to grant consent for the Graph delegated permissions listed below.',
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts53003')) {
    return {
      title: 'Conditional Access blocked',
      summary: 'Access is blocked by tenant Conditional Access policies.',
      description: 'Check your tenant Conditional Access policies or sign in from an approved device/location.',
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts50076') || lowered.includes('aadsts50079')) {
    return {
      title: 'Multi-factor authentication required',
      summary: 'Additional authentication steps are required by the tenant.',
      description: 'Complete the MFA prompt and retry the action.',
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts700082') || lowered.includes('token expired')) {
    return {
      title: 'Session expired',
      summary: 'Your session expired and tokens must be refreshed.',
      description: 'Sign out and sign back in to refresh tokens.',
      adminConsentUrl,
      details: message
    };
  }

  if (lowered.includes('aadsts50058') || lowered.includes('login_required')) {
    return {
      title: 'Login required',
      summary: 'You must sign in again to continue.',
      description: 'Sign in and retry the action.',
      adminConsentUrl,
      details: message
    };
  }

  if (isPopupOrCookieIssue(message, extractErrorCode(error))) {
    return {
      title: 'Popup or cookie blocked',
      summary: 'Your browser blocked the sign-in popup or third-party cookies.',
      description: 'Allow popups/cookies for this site or use a redirect-based sign-in.',
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

const SourceOption = ({ id, label, description, selected, disabled, onSelect }) => (
  <button
    type="button"
    className={`source-option ${selected ? 'active' : ''}`}
    onClick={() => onSelect(id)}
    disabled={disabled}
  >
    <strong>{label}</strong>
    <span>{description}</span>
  </button>
);

export default function App({ runtimeConfig, onSaveConfig, onResetConfig, onCopyConfigLink, loadWarnings = [] }) {
  const { instance, accounts } = useMsal();
  useEffect(() => {
    if (!instance.getActiveAccount() && accounts.length > 0) {
      instance.setActiveAccount(accounts[0]);
    }
  }, [accounts, instance]);

  const account = instance.getActiveAccount() ?? accounts[0];
  const tenantId = account?.tenantId ?? account?.idTokenClaims?.tid ?? null;

  const [draftConfig, setDraftConfig] = useState(runtimeConfig);
  const [sitCatalog, setSitCatalog] = useState([]);
  const [sitSelections, setSitSelections] = useState({});
  const [sitSearch, setSitSearch] = useState('');
  const [sitWarnings, setSitWarnings] = useState([]);
  const [sitDetectors, setSitDetectors] = useState([]);
  const [customDetectors, setCustomDetectors] = useState([]);
  const [showCustomRules, setShowCustomRules] = useState(false);
  const [sourceType, setSourceType] = useState('file');
  const [file, setFile] = useState(null);
  const [pasteText, setPasteText] = useState('');
  const [messages, setMessages] = useState([]);
  const [driveItems, setDriveItems] = useState([]);
  const [labels, setLabels] = useState([]);
  const [selectedMessageId, setSelectedMessageId] = useState('');
  const [selectedDriveItem, setSelectedDriveItem] = useState(null);
  const [extractedText, setExtractedText] = useState('');
  const [extractionMeta, setExtractionMeta] = useState(null);
  const [classificationResults, setClassificationResults] = useState([]);
  const [classificationWarnings, setClassificationWarnings] = useState([]);
  const [labelEvaluation, setLabelEvaluation] = useState(null);
  const [error, setError] = useState('');
  const [status, setStatus] = useState('');
  const [authHelp, setAuthHelp] = useState(null);
  const [loading, setLoading] = useState(false);
  const [consentLoading, setConsentLoading] = useState(false);

  useEffect(() => {
    setDraftConfig(runtimeConfig);
  }, [runtimeConfig]);

  const isAuthenticated = Boolean(account);
  const isConfigDirty = useMemo(
    () => JSON.stringify(draftConfig) !== JSON.stringify(runtimeConfig),
    [draftConfig, runtimeConfig]
  );

  const draftLoginScopes = useMemo(() => {
    if (Array.isArray(draftConfig.loginScopes)) {
      return draftConfig.loginScopes;
    }
    return String(draftConfig.loginScopes ?? '')
      .split(',')
      .map((scope) => scope.trim())
      .filter(Boolean);
  }, [draftConfig.loginScopes]);

  const invalidLoginScopes = useMemo(
    () => draftLoginScopes.filter((scope) => scope.includes('://') || scope.includes('.default')),
    [draftLoginScopes]
  );

  const graphClient = useMemo(
    () => createGraphClient({
      getToken: async (scopes) => {
        const activeAccount = instance.getActiveAccount() ?? accounts[0];
        if (!activeAccount) {
          throw new Error('No signed in account.');
        }
        try {
          const response = await instance.acquireTokenSilent({
            scopes,
            account: activeAccount
          });
          return response.accessToken;
        } catch (error) {
          if (error instanceof InteractionRequiredAuthError) {
            try {
              const response = await instance.acquireTokenPopup({
                scopes
              });
              if (response.account) {
                instance.setActiveAccount(response.account);
              }
              return response.accessToken;
            } catch (popupError) {
              const message = normalizeErrorMessage(popupError);
              if (shouldUseRedirect(message, extractErrorCode(popupError))) {
                await instance.acquireTokenRedirect({
                  scopes
                });
                return null;
              }
              throw popupError;
            }
          }
          throw error;
        }
      },
      graphBaseUrl: runtimeConfig.graphBaseUrl,
      defaultApiVersion: runtimeConfig.graphApiVersion
    }),
    [accounts, instance, runtimeConfig.graphApiVersion, runtimeConfig.graphBaseUrl]
  );

  const aggregatedResults = useMemo(
    () => aggregateClassificationResults(classificationResults),
    [classificationResults]
  );
  const sitPatternCount = useMemo(
    () => sitCatalog.reduce((total, sit) => total + (sit.patterns?.length ?? 0), 0),
    [sitCatalog]
  );
  const filteredSits = useMemo(() => {
    const query = sitSearch.trim().toLowerCase();
    if (!query) {
      return sitCatalog;
    }
    return sitCatalog.filter((sit) => (
      sit.name?.toLowerCase().includes(query)
      || sit.id?.toLowerCase().includes(query)
    ));
  }, [sitCatalog, sitSearch]);

  const handleLogin = async () => {
    if (!runtimeConfig.clientId) {
      setError('Add a client ID in Runtime Settings before signing in.');
      return;
    }
    setError('');
    setAuthHelp(null);

    try {
      const response = await instance.loginPopup(buildLoginRequest(runtimeConfig));
      if (response?.account) {
        instance.setActiveAccount(response.account);
      }
    } catch (popupError) {
      const message = normalizeErrorMessage(popupError);
      if (shouldUseRedirect(message, extractErrorCode(popupError))) {
        await instance.loginRedirect(buildLoginRequest(runtimeConfig));
        return;
      }
      const help = describeAuthError(popupError, tenantId, runtimeConfig);
      setAuthHelp(help);
      setError(help?.summary ?? popupError.message);
    }
  };

  const handleLogout = async () => {
    await instance.logoutPopup({
      account
    });
  };

  const handleSaveConfig = () => {
    onSaveConfig(draftConfig);
    setStatus('Runtime settings saved. Re-sign in if you changed tenant/client values.');
    setTimeout(() => setStatus(''), 4000);
  };

  const handleCopyLink = () => {
    onCopyConfigLink();
    setStatus('Config link copied to clipboard.');
    setTimeout(() => setStatus(''), 4000);
  };

  const handleExportConfig = () => {
    const blob = new Blob([JSON.stringify(draftConfig, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'runtime-config.json';
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const handleImportConfig = async (event) => {
    const fileInput = event.target.files?.[0];
    if (!fileInput) {
      return;
    }
    try {
      const text = await fileInput.text();
      const parsed = JSON.parse(text);
      setDraftConfig((prev) => ({ ...prev, ...parsed }));
      setStatus('Imported config. Click Save to apply.');
    } catch (error) {
      setError('Unable to parse runtime-config.json.');
    }
  };

  const handleImportRulePack = async (event) => {
    const fileInput = event.target.files?.[0];
    if (!fileInput) {
      return;
    }
    try {
      const buffer = await fileInput.arrayBuffer();
      const { sitCatalog: importedSits, warnings } = parseRulePackXml(buffer);
      const selection = {};
      importedSits.forEach((sit) => {
        selection[sit.id] = true;
      });
      setSitCatalog(importedSits);
      setSitSelections(selection);
      setSitWarnings(warnings ?? []);
      setSitDetectors([]);
      setClassificationResults([]);
      setLabelEvaluation(null);
      setClassificationWarnings([]);
      setStatus(`Imported ${importedSits.length} SITs from rule pack.`);
      setTimeout(() => setStatus(''), 4000);
    } catch (error) {
      setError('Unable to parse rule pack XML.');
    }
  };

  const toggleSitSelection = (sitId) => {
    setSitSelections((prev) => ({
      ...prev,
      [sitId]: !prev[sitId]
    }));
  };

  const selectAllSits = () => {
    const selection = {};
    sitCatalog.forEach((sit) => {
      selection[sit.id] = true;
    });
    setSitSelections(selection);
  };

  const clearSitSelection = () => {
    setSitSelections({});
  };

  const buildDetectorsFromSits = (sits) => {
    const detectors = [];
    sits.forEach((sit) => {
      sit.patterns.forEach((pattern) => {
        if (!pattern.nodes?.length) {
          return;
        }
        detectors.push({
          id: `${sit.id}-${pattern.id ?? sit.id}`,
          sitId: sit.id,
          name: sit.name,
          description: sit.description,
          pattern,
          confidence: pattern.confidence ?? 50,
          source: 'rulepack'
        });
      });
    });
    return detectors;
  };

  const applySelectedSits = ({ append = false } = {}) => {
    const selectedSits = sitCatalog.filter((sit) => sitSelections[sit.id]);
    if (!selectedSits.length) {
      setError('Select at least one SIT to load.');
      return;
    }

    const nextDetectors = buildDetectorsFromSits(selectedSits);
    if (!nextDetectors.length) {
      setError('No valid patterns were found in the selected SITs.');
      return;
    }

    setSitDetectors((prev) => (append ? [...prev, ...nextDetectors] : nextDetectors));
    setClassificationResults([]);
    setLabelEvaluation(null);
    setClassificationWarnings([]);
    const action = append ? 'Appended' : 'Loaded';
    setStatus(`${action} ${selectedSits.length} SITs (${nextDetectors.length} patterns).`);
    setTimeout(() => setStatus(''), 4000);
  };

  const handleLoadSampleDetectors = () => {
    const samples = getSampleDetectors().map((detector) => ({
      ...detector,
      source: 'sample'
    }));
    setCustomDetectors(samples);
    setShowCustomRules(true);
  };

  const handleGrantScopes = async (scopes) => {
    const mergedScopes = buildScopeList([runtimeConfig.profileScope, ...scopes]);
    if (!mergedScopes.length) {
      setError('No scopes configured for this feature.');
      return;
    }
    setConsentLoading(true);
    setError('');
    setAuthHelp(null);
    try {
      await graphClient.request({
        path: '/me',
        scopes: mergedScopes
      });
      setStatus('Permissions granted successfully.');
    } catch (err) {
      const help = describeAuthError(err, tenantId, runtimeConfig);
      setAuthHelp(help);
      setError(help?.summary ?? err.message);
    } finally {
      setConsentLoading(false);
    }
  };

  const loadMessages = async () => {
    setLoading(true);
    setError('');
    setAuthHelp(null);
    try {
      const items = await graphClient.getAllPages({
        path: '/me/messages?$top=10&$select=id,subject,receivedDateTime,from,bodyPreview',
        scopes: buildScopeList([runtimeConfig.mailReadScope])
      });
      setMessages(items);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const loadDriveItems = async () => {
    setLoading(true);
    setError('');
    setAuthHelp(null);
    try {
      const items = await graphClient.request({
        path: '/me/drive/recent?$top=10',
        scopes: buildScopeList([runtimeConfig.filesReadScope, runtimeConfig.sitesReadScope])
      });
      const files = (items?.value ?? []).filter((item) => item.file);
      setDriveItems(files);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const loadSensitivityLabels = async () => {
    setLoading(true);
    setError('');
    setAuthHelp(null);
    try {
      const response = await graphClient.request({
        path: '/me/security/informationProtection/sensitivityLabels?$top=50',
        apiVersion: runtimeConfig.infoProtectionApiVersion,
        scopes: buildScopeList([runtimeConfig.labelsReadScope])
      });
      setLabels(response?.value ?? []);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleExtract = async () => {
    setLoading(true);
    setError('');
    setAuthHelp(null);
    setLabelEvaluation(null);
    setClassificationResults([]);
    setClassificationWarnings([]);

    try {
      if (sourceType === 'file') {
        if (!file) {
          throw new Error('Select a file to extract.');
        }
        const text = await extractTextFromFile(file);
        setExtractedText(text);
        setExtractionMeta({
          source: 'Local file',
          name: file.name,
          length: text.length
        });
        return;
      }

      if (sourceType === 'paste') {
        if (!pasteText.trim()) {
          throw new Error('Paste or type text to extract.');
        }
        setExtractedText(pasteText.trim());
        setExtractionMeta({
          source: 'Pasted text',
          name: 'Pasted text',
          length: pasteText.trim().length
        });
        return;
      }

      if (sourceType === 'mail') {
        if (!selectedMessageId) {
          throw new Error('Select a message.');
        }
        const message = await graphClient.request({
          path: `/me/messages/${selectedMessageId}?$select=subject,receivedDateTime,from,body`,
          scopes: buildScopeList([runtimeConfig.mailReadScope])
        });
        const bodyText = extractTextFromHtml(message?.body?.content ?? '');
        setExtractedText(bodyText);
        setExtractionMeta({
          source: 'Outlook message',
          name: message?.subject ?? 'Message',
          length: bodyText.length
        });
        return;
      }

      if (sourceType === 'drive') {
        if (!selectedDriveItem) {
          throw new Error('Select a OneDrive file.');
        }
        const buffer = await graphClient.downloadContent({
          path: `/me/drive/items/${selectedDriveItem.id}/content`,
          scopes: buildScopeList([runtimeConfig.filesReadScope, runtimeConfig.sitesReadScope])
        });
        const text = await extractTextFromBuffer({ buffer, name: selectedDriveItem.name });
        setExtractedText(text);
        setExtractionMeta({
          source: 'OneDrive file',
          name: selectedDriveItem.name,
          length: text.length
        });
        return;
      }

      throw new Error('Select a valid data source.');
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleClassify = () => {
    if (!extractedText) {
      setError('Extract text before running classification.');
      return;
    }
    if (!sitDetectors.length && !customDetectors.length) {
      setError('Load SITs or add custom rules before running classification.');
      return;
    }
    setError('');
    const invalidCustomDetectors = findInvalidDetectors(customDetectors);
    if (invalidCustomDetectors.length > 0) {
      setClassificationWarnings([
        `${invalidCustomDetectors.length} custom rule${invalidCustomDetectors.length === 1 ? '' : 's'} were skipped because their regex patterns are invalid.`
      ]);
    } else {
      setClassificationWarnings([]);
    }
    const validCustomDetectors = customDetectors.filter((detector) => !invalidCustomDetectors.includes(detector));
    const sitResults = evaluateSitDetectors(extractedText, sitDetectors);
    const customResults = detectSensitiveInfo(extractedText, validCustomDetectors);
    const results = [...sitResults, ...customResults];
    if (!results.length) {
      setError('No matches were found for the current rules.');
    }
    setClassificationResults(results);
    setLabelEvaluation(null);
  };

  const handleEvaluateLabels = async () => {
    const graphResults = buildGraphClassificationResults(classificationResults);
    if (!graphResults.length) {
      setError('Load SITs (with sensitiveTypeId values) or update your custom rules before evaluating labels.');
      return;
    }
    setLoading(true);
    setError('');
    setAuthHelp(null);

    try {
      const payload = {
        contentInfo: {
          '@odata.type': '#microsoft.graph.security.contentInfo',
          'format@odata.type': '#microsoft.graph.security.contentFormat',
          format: 'default',
          contentFormat: 'File',
          identifier: extractionMeta?.name ?? 'extracted.txt',
          'state@odata.type': '#microsoft.graph.security.contentState',
          state: 'rest',
          metadata: []
        },
        classificationResults: graphResults
      };

      const response = await graphClient.request({
        path: '/me/security/informationProtection/sensitivityLabels/evaluateClassificationResults',
        method: 'POST',
        apiVersion: runtimeConfig.infoProtectionApiVersion,
        scopes: buildScopeList([runtimeConfig.labelsReadScope]),
        body: payload
      });
      setLabelEvaluation(response);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const addCustomDetector = () => {
    setCustomDetectors((prev) => ([
      ...prev,
      {
        id: `custom-${Date.now()}`,
        name: 'Custom rule',
        description: 'Custom regex pattern.',
        pattern: '',
        minCount: 1,
        confidence: 50,
        sensitiveTypeId: '',
        source: 'manual'
      }
    ]));
    setShowCustomRules(true);
  };

  const updateCustomDetector = (id, changes) => {
    setCustomDetectors((prev) => prev.map((detector) => (
      detector.id === id ? { ...detector, ...changes } : detector
    )));
  };

  const removeCustomDetector = (id) => {
    setCustomDetectors((prev) => prev.filter((detector) => detector.id !== id));
  };

  const scopeSummary = {
    login: summarizeScopeSet(runtimeConfig.loginScopes),
    profile: summarizeScopeSet(buildScopeList([runtimeConfig.profileScope])),
    mail: summarizeScopeSet(buildScopeList([runtimeConfig.mailReadScope])),
    files: summarizeScopeSet(buildScopeList([runtimeConfig.filesReadScope, runtimeConfig.sitesReadScope])),
    labels: summarizeScopeSet(buildScopeList([runtimeConfig.labelsReadScope]))
  };

  const adminConsentUrl = buildGraphAdminConsentUrl(runtimeConfig, tenantId);

  return (
    <div className="app">
      <header className="hero">
        <div>
          <p className="eyebrow">Microsoft Graph + Purview</p>
          <h1>Browser-Only Extraction & Classification</h1>
          <p>
            Replicate Test-TextExtraction and Test-DataClassification outcomes using in-browser processing and Microsoft Graph.
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
            <button type="button" onClick={handleLogin} disabled={!runtimeConfig.clientId}>
              Sign in with Microsoft 365
            </button>
          )}
        </div>
      </header>

      <main>
        <section className="card">
          <h2>Runtime Settings</h2>
          <p className="helper">
            Configure tenant-specific settings at runtime. Values are stored in your browser (localStorage) and can
            also be provided via URL parameters.
          </p>
          {loadWarnings.length > 0 && (
            <p className="warning">{loadWarnings.join(' ')}</p>
          )}
          {!runtimeConfig.clientId && (
            <p className="warning">Client ID is required to sign in.</p>
          )}
          {invalidLoginScopes.length > 0 && (
            <p className="warning">
              Login scopes must only include openid/profile/email. Remove: {invalidLoginScopes.join(', ')}
            </p>
          )}
          <div className="config-grid">
            <div className="field">
              <label>Client (Application) ID</label>
              <input
                type="text"
                value={draftConfig.clientId ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, clientId: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Authority Host</label>
              <input
                type="text"
                value={draftConfig.authorityHost ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, authorityHost: event.target.value }))}
              />
              <span className="hint">Example: https://login.microsoftonline.com</span>
            </div>
            <div className="field">
              <label>Tenant</label>
              <input
                type="text"
                value={draftConfig.authorityTenant ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, authorityTenant: event.target.value }))}
              />
              <span className="hint">Use organizations, common, or a tenant ID.</span>
            </div>
            <div className="field">
              <label>Redirect URI</label>
              <input
                type="text"
                value={draftConfig.redirectUri ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, redirectUri: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Microsoft Graph Base URL</label>
              <input
                type="text"
                value={draftConfig.graphBaseUrl ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, graphBaseUrl: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Graph API Version</label>
              <input
                type="text"
                value={draftConfig.graphApiVersion ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, graphApiVersion: event.target.value }))}
              />
              <span className="hint">Default: v1.0</span>
            </div>
            <div className="field">
              <label>Info Protection API Version</label>
              <input
                type="text"
                value={draftConfig.infoProtectionApiVersion ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, infoProtectionApiVersion: event.target.value }))}
              />
              <span className="hint">Default: beta</span>
            </div>
            <div className="field">
              <label>Login Scopes</label>
              <input
                type="text"
                value={Array.isArray(draftConfig.loginScopes) ? draftConfig.loginScopes.join(',') : draftConfig.loginScopes ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, loginScopes: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>User.Read Scope</label>
              <input
                type="text"
                value={draftConfig.profileScope ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, profileScope: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Mail.Read Scope</label>
              <input
                type="text"
                value={draftConfig.mailReadScope ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, mailReadScope: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Files.Read Scope</label>
              <input
                type="text"
                value={draftConfig.filesReadScope ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, filesReadScope: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>Sites.Read.All Scope (optional)</label>
              <input
                type="text"
                value={draftConfig.sitesReadScope ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, sitesReadScope: event.target.value }))}
              />
            </div>
            <div className="field">
              <label>InformationProtectionPolicy.Read Scope</label>
              <input
                type="text"
                value={draftConfig.labelsReadScope ?? ''}
                onChange={(event) => setDraftConfig((prev) => ({ ...prev, labelsReadScope: event.target.value }))}
              />
            </div>
          </div>
          <div className="actions">
            <button type="button" onClick={handleSaveConfig} disabled={!isConfigDirty}>
              Save runtime settings
            </button>
            <button type="button" className="secondary" onClick={onResetConfig}>
              Reset settings
            </button>
            <button type="button" className="secondary" onClick={handleCopyLink}>
              Copy config link
            </button>
            <button type="button" className="secondary" onClick={handleExportConfig}>
              Export JSON
            </button>
            <label className="secondary file-button">
              Import JSON
              <input type="file" accept="application/json" onChange={handleImportConfig} />
            </label>
          </div>
          {status && <p className="status">{status}</p>}
        </section>

        <section className="card">
          <h2>Permissions & Consent</h2>
          <p className="helper">
            This app uses delegated permissions only. Login scopes are requested during sign-in, and Graph scopes are
            requested on-demand.
          </p>
          <div className="permission-list">
            <div>
              <strong>Login scopes</strong>
              <p>{formatScopeList(runtimeConfig.loginScopes)} — basic sign-in.</p>
            </div>
            <div>
              <strong>Profile (required)</strong>
              <p>{scopeSummary.profile} — required to validate Graph access.</p>
            </div>
            <div>
              <strong>Mail (optional)</strong>
              <p>{scopeSummary.mail} — required to read Outlook message bodies.</p>
            </div>
            <div>
              <strong>Files (optional)</strong>
              <p>{scopeSummary.files} — required to read OneDrive/SharePoint file content.</p>
            </div>
            <div>
              <strong>Information Protection (optional)</strong>
              <p>{scopeSummary.labels} — required to list labels and evaluate classification results.</p>
            </div>
          </div>
          <div className="actions">
            <button
              type="button"
              onClick={() => handleGrantScopes(buildScopeList([
                runtimeConfig.mailReadScope,
                runtimeConfig.filesReadScope,
                runtimeConfig.sitesReadScope,
                runtimeConfig.labelsReadScope
              ]))}
              disabled={!isAuthenticated || consentLoading}
            >
              Grant selected Graph scopes
            </button>
            {adminConsentUrl && (
              <button
                type="button"
                className="secondary"
                onClick={() => window.open(adminConsentUrl, '_blank', 'noopener,noreferrer')}
              >
                Admin consent link
              </button>
            )}
          </div>
          {consentLoading && <p className="status">Requesting permissions...</p>}
        </section>

        <section className="card">
          <h2>Data Source</h2>
          <div className="source-grid">
            <SourceOption
              id="file"
              label="Local file"
              description="PDF, DOCX, TXT, CSV, MD, JSON, EML (client-side extraction)"
              selected={sourceType === 'file'}
              onSelect={setSourceType}
            />
            <SourceOption
              id="paste"
              label="Paste text"
              description="Paste or type text directly"
              selected={sourceType === 'paste'}
              onSelect={setSourceType}
            />
            <SourceOption
              id="mail"
              label="Outlook email (Graph)"
              description="Load recent messages via Mail.Read"
              selected={sourceType === 'mail'}
              onSelect={setSourceType}
              disabled={!runtimeConfig.mailReadScope}
            />
            <SourceOption
              id="drive"
              label="OneDrive recent file (Graph)"
              description="Load recent files via Files.Read"
              selected={sourceType === 'drive'}
              onSelect={setSourceType}
              disabled={!runtimeConfig.filesReadScope}
            />
          </div>

          {sourceType === 'file' && (
            <div className="field">
              <label htmlFor="file">Local file</label>
              <input
                id="file"
                type="file"
                accept=".pdf,.docx,.txt,.csv,.md,.json,.eml"
                onChange={(event) => setFile(event.target.files?.[0] ?? null)}
              />
              <p className="helper">
                Supports PDF, DOCX, TXT, CSV, MD, JSON, EML today; attachments, images, and archives will appear after the next extractor update.
              </p>
            </div>
          )}

          {sourceType === 'paste' && (
            <div className="field">
              <label>Paste text</label>
              <textarea
                rows={6}
                value={pasteText}
                onChange={(event) => setPasteText(event.target.value)}
              />
            </div>
          )}

          {sourceType === 'mail' && (
            <div className="field">
              <div className="actions">
                <button type="button" className="secondary" onClick={loadMessages} disabled={!isAuthenticated || loading}>
                  Load recent messages
                </button>
              </div>
              <select
                value={selectedMessageId}
                onChange={(event) => setSelectedMessageId(event.target.value)}
              >
                <option value="">Select a message</option>
                {messages.map((message) => (
                  <option key={message.id} value={message.id}>
                    {message.subject || '(no subject)'}
                  </option>
                ))}
              </select>
            </div>
          )}

          {sourceType === 'drive' && (
            <div className="field">
              <div className="actions">
                <button type="button" className="secondary" onClick={loadDriveItems} disabled={!isAuthenticated || loading}>
                  Load recent files
                </button>
              </div>
              <select
                value={selectedDriveItem?.id ?? ''}
                onChange={(event) => {
                  const item = driveItems.find((entry) => entry.id === event.target.value);
                  setSelectedDriveItem(item ?? null);
                }}
              >
                <option value="">Select a file</option>
                {driveItems.map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="actions">
            <button type="button" onClick={handleExtract} disabled={loading}>
              Run text extraction
            </button>
          </div>
        </section>

        {extractionMeta && (
          <section className="card">
            <h2>Extraction Results</h2>
            <div className="meta-grid">
              <div>
                <strong>Source</strong>
                <p>{extractionMeta.source}</p>
              </div>
              <div>
                <strong>Identifier</strong>
                <p>{extractionMeta.name}</p>
              </div>
              <div>
                <strong>Characters</strong>
                <p>{extractionMeta.length}</p>
              </div>
            </div>
            <pre className="text-preview">{extractedText.slice(0, 5000)}</pre>
          </section>
        )}

        <section className="card">
          <h2>Sensitive Information Types (SITs)</h2>
          <p className="helper">
            Load SIT definitions from your tenant by importing an XML rule pack. These rules can be applied to the extracted
            text before evaluating labels.
          </p>
          <div className="sit-actions">
            <label className="secondary file-button">
              Import rule pack XML
              <input type="file" accept=".xml" onChange={handleImportRulePack} />
            </label>
            <button type="button" className="secondary" onClick={selectAllSits} disabled={!sitCatalog.length}>
              Select all
            </button>
            <button type="button" className="secondary" onClick={clearSitSelection} disabled={!sitCatalog.length}>
              Select none
            </button>
            <button type="button" onClick={() => applySelectedSits({ append: false })} disabled={!sitCatalog.length}>
              Load selected SITs
            </button>
            <button type="button" className="secondary" onClick={() => applySelectedSits({ append: true })} disabled={!sitCatalog.length}>
              Append to rules
            </button>
          </div>
          {sitCatalog.length === 0 && (
            <div className="helper">
              <p>Export a rule pack XML (PowerShell) and import it here:</p>
              <pre>
                {`$rulePack = Get-DlpSensitiveInformationTypeRulePackage -Identity "Microsoft Rule Package"
[System.IO.File]::WriteAllBytes("./rulepack.xml", $rulePack.SerializedClassificationRuleCollection)`}
              </pre>
            </div>
          )}
          {sitWarnings.length > 0 && (
            <div className="warning">
              {sitWarnings.map((warning) => (
                <p key={warning}>{warning}</p>
              ))}
            </div>
          )}
          {sitCatalog.length > 0 && (
            <>
              <div className="field">
                <label>Search SITs</label>
                <input
                  type="text"
                  value={sitSearch}
                  onChange={(event) => setSitSearch(event.target.value)}
                  placeholder="Filter by name or ID"
                />
                <span className="hint">
                  Loaded {sitCatalog.length} SITs / {sitPatternCount} regex patterns.
                </span>
              </div>
              <div className="sit-list">
                {filteredSits.map((sit) => (
                  <label key={sit.id} className="sit-item">
                    <input
                      type="checkbox"
                      checked={Boolean(sitSelections[sit.id])}
                      onChange={() => toggleSitSelection(sit.id)}
                    />
                    <div>
                      <strong>{sit.name}</strong>
                      <span>{sit.id}</span>
                      {sit.description && <em>{sit.description}</em>}
                    </div>
                  </label>
                ))}
              </div>
            </>
          )}
        </section>

        <section className="card">
          <h2>Data Classification (client-side)</h2>
          <p className="helper">
            Classification rules run locally in the browser. Load SITs above or add custom rules to evaluate extracted text.
          </p>
          <div className="meta-grid">
            <div>
              <strong>Rule sources</strong>
              <p>{sitDetectors.length} from SITs, {customDetectors.length} custom.</p>
            </div>
          </div>
          <div className="actions">
            <button type="button" onClick={handleClassify}>
              Run classification
            </button>
            <button type="button" className="secondary" onClick={() => setShowCustomRules((prev) => !prev)}>
              {showCustomRules ? 'Hide custom rules' : 'Show custom rules'}
            </button>
            <button type="button" className="secondary" onClick={handleLoadSampleDetectors}>
              Load sample rules
            </button>
            <button type="button" className="secondary" onClick={addCustomDetector}>
              Add custom rule
            </button>
          </div>
          {classificationWarnings.length > 0 && (
            <div className="warning">
              {classificationWarnings.map((warning) => (
                <p key={warning}>{warning}</p>
              ))}
            </div>
          )}
          {(showCustomRules || customDetectors.length > 0) && (
            <div className="detector-list">
              {customDetectors.map((detector) => (
                <div key={detector.id} className="detector-card">
                  <div className="field">
                    <label>Name</label>
                    <input
                      type="text"
                      value={detector.name}
                      onChange={(event) => updateCustomDetector(detector.id, { name: event.target.value })}
                    />
                  </div>
                  <div className="field">
                    <label>Regex Pattern</label>
                    <input
                      type="text"
                      value={detector.pattern}
                      onChange={(event) => updateCustomDetector(detector.id, { pattern: event.target.value })}
                    />
                  </div>
                  <div className="field">
                    <label>Min Count</label>
                    <input
                      type="number"
                      min={1}
                      value={detector.minCount}
                      onChange={(event) => updateCustomDetector(detector.id, { minCount: Number(event.target.value) })}
                    />
                  </div>
                  <div className="field">
                    <label>Confidence</label>
                    <input
                      type="number"
                      min={0}
                      max={99}
                      value={detector.confidence}
                      onChange={(event) => updateCustomDetector(detector.id, { confidence: Number(event.target.value) })}
                    />
                  </div>
                  <div className="field">
                    <label>SensitiveTypeId (optional)</label>
                    <input
                      type="text"
                      value={detector.sensitiveTypeId}
                      onChange={(event) => updateCustomDetector(detector.id, { sensitiveTypeId: event.target.value })}
                    />
                  </div>
                  <button type="button" className="secondary" onClick={() => removeCustomDetector(detector.id)}>
                    Remove rule
                  </button>
                </div>
              ))}
            </div>
          )}
        </section>

        {aggregatedResults.length > 0 && (
          <section className="card">
            <h2>Classification Results</h2>
            <div className="results-grid">
              {aggregatedResults.map((result) => (
                <article key={result.sensitiveTypeId || result.id}>
                  <header>
                    <strong>{result.name}</strong>
                    <span>Count: {result.count}</span>
                  </header>
                  <p>Confidence: {result.confidence}</p>
                  {result.sensitiveTypeId && <p>SIT ID: {result.sensitiveTypeId}</p>}
                  {result.samples?.length > 0 && (
                    <pre>{result.samples.join('\n')}</pre>
                  )}
                </article>
              ))}
            </div>
            <div className="actions">
              <button type="button" className="secondary" onClick={handleEvaluateLabels} disabled={!isAuthenticated || loading}>
                Evaluate labels (Graph)
              </button>
              <button type="button" className="secondary" onClick={loadSensitivityLabels} disabled={!isAuthenticated || loading}>
                Load sensitivity labels
              </button>
            </div>
          </section>
        )}

        {labels.length > 0 && (
          <section className="card">
            <h2>Sensitivity Labels (Graph)</h2>
            <div className="results-grid">
              {labels.map((label) => (
                <article key={label.id}>
                  <header>
                    <strong>{label.name}</strong>
                    <span>{label.id}</span>
                  </header>
                  <pre>{formatJson(label)}</pre>
                </article>
              ))}
            </div>
          </section>
        )}

        {labelEvaluation && (
          <ResultCard title="Label Evaluation Response" content={labelEvaluation} />
        )}

        {authHelp && (
          <section className="card notice">
            <h3>{authHelp.title}</h3>
            <p>{authHelp.description}</p>
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

        {error && <p className="error">{error}</p>}
      </main>

      <footer>
        <p>
          This experience runs entirely in the browser. No server-side token exchange or secrets are required.
        </p>
      </footer>
    </div>
  );
}
