const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const parseRetryAfter = (value) => {
  if (!value) {
    return null;
  }
  const seconds = Number(value);
  if (!Number.isNaN(seconds)) {
    return seconds * 1000;
  }
  const date = Date.parse(value);
  if (!Number.isNaN(date)) {
    const delay = date - Date.now();
    return delay > 0 ? delay : null;
  }
  return null;
};

const buildRequestId = () => {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
};

const defaultRetryStatuses = new Set([429, 500, 502, 503, 504]);

export const createGraphClient = ({ getToken, graphBaseUrl, defaultApiVersion = 'v1.0' }) => {
  const request = async ({
    path,
    method = 'GET',
    apiVersion = defaultApiVersion,
    scopes = [],
    headers = {},
    body,
    maxRetries = 3
  }) => {
    const url = path.startsWith('http')
      ? path
      : `${graphBaseUrl.replace(/\/+$/, '')}/${apiVersion}${path}`;

    const token = scopes.length ? await getToken(scopes) : null;
    if (scopes.length && !token) {
      throw new Error('No access token available for Microsoft Graph request.');
    }

    let lastError;
    for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
      try {
        const requestHeaders = {
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
          'client-request-id': buildRequestId(),
          ...headers
        };
        if (body) {
          requestHeaders['Content-Type'] = 'application/json';
        }

        const response = await fetch(url, {
          method,
          headers: requestHeaders,
          body: body ? JSON.stringify(body) : undefined
        });

        if (!response.ok) {
          const retryAfterMs = parseRetryAfter(response.headers.get('Retry-After'));
          if (defaultRetryStatuses.has(response.status) && attempt < maxRetries) {
            const backoff = retryAfterMs ?? Math.min(1000 * 2 ** attempt, 8000);
            await sleep(backoff + Math.floor(Math.random() * 250));
            continue;
          }

          let payload = null;
          try {
            payload = await response.json();
          } catch (error) {
            // ignore parse errors
          }
          const errorMessage = payload?.error?.message || `Graph request failed (${response.status}).`;
          const errorCode = payload?.error?.code;
          const err = new Error(errorMessage);
          err.status = response.status;
          err.code = errorCode;
          err.payload = payload;
          throw err;
        }

        if (response.status === 204) {
          return null;
        }

        const contentType = response.headers.get('content-type') ?? '';
        if (contentType.includes('application/json')) {
          return response.json();
        }
        return response;
      } catch (error) {
        lastError = error;
        if (attempt >= maxRetries) {
          throw error;
        }
        await sleep(Math.min(500 * 2 ** attempt, 4000));
      }
    }

    throw lastError ?? new Error('Graph request failed.');
  };

  const getAllPages = async ({ path, scopes = [], apiVersion = defaultApiVersion, maxPages = 10 }) => {
    let next = path;
    const items = [];
    let pageCount = 0;

    while (next && pageCount < maxPages) {
      const response = await request({ path: next, scopes, apiVersion });
      if (response?.value) {
        items.push(...response.value);
      }
      next = response?.['@odata.nextLink'] ?? null;
      pageCount += 1;
    }

    return items;
  };

  const downloadContent = async ({ path, scopes = [], apiVersion = defaultApiVersion }) => {
    const token = scopes.length ? await getToken(scopes) : null;
    if (scopes.length && !token) {
      throw new Error('No access token available for Microsoft Graph request.');
    }
    const url = `${graphBaseUrl.replace(/\/+$/, '')}/${apiVersion}${path}`;
    const response = await fetch(url, {
      headers: {
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        'client-request-id': buildRequestId()
      }
    });
    if (!response.ok) {
      const errorMessage = `Graph content download failed (${response.status}).`;
      const err = new Error(errorMessage);
      err.status = response.status;
      throw err;
    }
    return response.arrayBuffer();
  };

  return {
    request,
    getAllPages,
    downloadContent
  };
};
