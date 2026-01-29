import jwt from 'jsonwebtoken';
import jwksRsa from 'jwks-rsa';
import { config } from './config.js';
import { logger } from './logger.js';

const jwksClients = new Map();

const getJwksClient = (authorityHost, tenantId) => {
  const key = `${authorityHost}/${tenantId}`;
  if (!jwksClients.has(key)) {
    jwksClients.set(key, jwksRsa({
      cache: true,
      rateLimit: true,
      jwksRequestsPerMinute: 5,
      jwksUri: `${authorityHost}/${tenantId}/discovery/v2.0/keys`
    }));
  }
  return jwksClients.get(key);
};

const extractBearerToken = (req) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith('Bearer ')) {
    return null;
  }
  return authHeader.substring('Bearer '.length).trim();
};

const resolveIssuerContext = (claims) => {
  const issuer = claims?.iss;
  const tenantId = claims?.tid;

  if (!issuer || !tenantId) {
    throw new Error('Token missing issuer or tenant.');
  }

  let issuerUrl;
  try {
    issuerUrl = new URL(issuer);
  } catch (error) {
    throw new Error('Token issuer is invalid.');
  }

  const authorityHost = issuerUrl.origin;
  if (!config.allowedAuthorityHosts.includes(authorityHost)) {
    throw new Error(`Authority host not allowed: ${authorityHost}`);
  }

  const expectedIssuer = `${authorityHost}/${tenantId}/v2.0`;
  if (issuer !== expectedIssuer) {
    throw new Error('Token issuer does not match expected tenant issuer.');
  }

  return { authorityHost, tenantId, expectedIssuer };
};

const verifyAccessToken = async (token, options = {}) => {
  const decoded = jwt.decode(token, { complete: true });
  if (!decoded?.payload) {
    throw new Error('Invalid access token.');
  }

  const { authorityHost, tenantId, expectedIssuer } = resolveIssuerContext(decoded.payload);
  if (options.expectedTenantId && options.expectedTenantId !== tenantId) {
    throw new Error('Token tenant mismatch.');
  }

  const client = getJwksClient(authorityHost, tenantId);
  const getKey = (header, callback) => {
    client.getSigningKey(header.kid, (err, key) => {
      if (err) {
        callback(err);
        return;
      }
      callback(null, key.getPublicKey());
    });
  };

  const verifyOptions = {
    algorithms: ['RS256'],
    issuer: expectedIssuer,
    ...(options.expectedAudiences?.length ? { audience: options.expectedAudiences } : {})
  };

  const verified = await new Promise((resolve, reject) => {
    jwt.verify(token, getKey, verifyOptions, (err, payload) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(payload);
    });
  });

  return {
    claims: verified,
    tenantId,
    issuer: expectedIssuer,
    authorityHost
  };
};

const audienceMatches = (audience, allowedAudiences) => {
  if (!audience) {
    return false;
  }
  if (Array.isArray(audience)) {
    return audience.some((value) => allowedAudiences.includes(value));
  }
  return allowedAudiences.includes(audience);
};

export const tokenAudiences = {
  exchange: config.exchangeAudiences,
  compliance: config.complianceAudiences,
  classification: Array.from(new Set([...config.exchangeAudiences, ...config.complianceAudiences]))
};

export const authenticate = (expectedAudiences = tokenAudiences.classification) => async (req, res, next) => {
  const token = extractBearerToken(req);
  if (!token) {
    res.status(401).json({ error: 'Missing bearer token.' });
    return;
  }

  try {
    const { claims, tenantId, issuer, authorityHost } = await verifyAccessToken(token, { expectedAudiences });

    if (config.allowedTenants.length > 0 && !config.allowedTenants.includes(tenantId)) {
      res.status(403).json({ error: 'Tenant not allowed.' });
      return;
    }

    req.auth = {
      tenantId,
      issuer,
      authorityHost,
      userPrincipalName: claims.preferred_username ?? claims.upn ?? claims.email,
      token,
      audience: claims.aud
    };

    logger.info('auth_success', { user: req.auth.userPrincipalName, tenant: tenantId });
    next();
  } catch (error) {
    logger.warn('auth_failed', { error: error.message });
    res.status(401).json({ error: 'Invalid token.' });
  }
};

export const verifySupplementalToken = async (token, options) => verifyAccessToken(token, options);

export const isAudienceAllowed = audienceMatches;
