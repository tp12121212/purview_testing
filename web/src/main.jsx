import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { MsalProvider } from '@azure/msal-react';
import { PublicClientApplication } from '@azure/msal-browser';
import App from './App.jsx';
import { buildMsalConfig } from './authConfig.js';
import {
  DEFAULT_RUNTIME_CONFIG,
  buildConfigQueryString,
  loadRuntimeConfig,
  normalizeRuntimeConfig,
  resetRuntimeConfig,
  saveRuntimeConfig
} from './runtimeConfig.js';
import './styles.css';

const LoadingScreen = ({ message }) => (
  <div className="app app-loading">
    <p>{message}</p>
  </div>
);

const Root = () => {
  const [runtimeConfig, setRuntimeConfig] = useState(null);
  const [warnings, setWarnings] = useState([]);

  useEffect(() => {
    let isMounted = true;
    loadRuntimeConfig()
      .then(({ config, warnings: loadWarnings }) => {
        if (!isMounted) {
          return;
        }
        setRuntimeConfig(config);
        setWarnings(loadWarnings ?? []);
      })
      .catch(() => {
        if (!isMounted) {
          return;
        }
        setRuntimeConfig(DEFAULT_RUNTIME_CONFIG);
        setWarnings(['Failed to load runtime config. Using defaults.']);
      });
    return () => {
      isMounted = false;
    };
  }, []);

  const msalInstance = useMemo(() => {
    if (!runtimeConfig) {
      return null;
    }
    return new PublicClientApplication(buildMsalConfig(runtimeConfig));
  }, [runtimeConfig]);

  const handleSaveConfig = (nextConfig) => {
    const normalized = normalizeRuntimeConfig(nextConfig);
    saveRuntimeConfig(normalized);
    setRuntimeConfig(normalized);
  };

  const handleResetConfig = () => {
    resetRuntimeConfig();
    window.location.reload();
  };

  const handleCopyConfigLink = async () => {
    if (!runtimeConfig) {
      return;
    }
    const link = `${window.location.origin}${window.location.pathname}${buildConfigQueryString(runtimeConfig)}`;
    if (navigator.clipboard?.writeText) {
      try {
        await navigator.clipboard.writeText(link);
        return;
      } catch (error) {
        // Fall back to a prompt if clipboard access fails.
      }
    }
    window.prompt('Copy config link', link);
  };

  if (!runtimeConfig || !msalInstance) {
    return <LoadingScreen message="Loading runtime configuration..." />;
  }

  return (
    <MsalProvider
      instance={msalInstance}
      key={`${runtimeConfig.clientId}-${runtimeConfig.authorityHost}-${runtimeConfig.authorityTenant}`}
    >
      <App
        runtimeConfig={runtimeConfig}
        onSaveConfig={handleSaveConfig}
        onResetConfig={handleResetConfig}
        onCopyConfigLink={handleCopyConfigLink}
        loadWarnings={warnings}
      />
    </MsalProvider>
  );
};

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
