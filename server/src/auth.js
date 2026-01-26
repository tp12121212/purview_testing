import jwt from 'jsonwebtoken';
import jwksRsa from 'jwks-rsa';
import { config } from './config.js';
import { logger } from './logger.js';

const jwksClient = jwksRsa({
  cache: true,
  rateLimit: true,
  jwksRequestsPerMinute: 5,
  jwksUri: `${config.authorityHost}/common/discovery/v2.0/keys`
});

const getKey = (header, callback) => {
  jwksClient.getSigningKey(header.kid, (err, key) => {
    if (err) {
      callback(err);
      return;
    }
    const signingKey = key.getPublicKey();
    callback(null, signingKey);
  });
};

const issuerForTenant = (tenantId) => `${config.authorityHost}/${tenantId}/v2.0`;

export const authenticate = (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith('Bearer ')) {
    res.status(401).json({ error: 'Missing bearer token.' });
    return;
  }

  const token = authHeader.substring('Bearer '.length);

  jwt.verify(token, getKey, { algorithms: ['RS256'], audience: config.clientId }, (err, decoded) => {
    if (err) {
      logger.warn('auth_failed', { error: err.message });
      res.status(401).json({ error: 'Invalid token.' });
      return;
    }

    const tenantId = decoded.tid;
    if (!tenantId) {
      res.status(401).json({ error: 'Token missing tenant id.' });
      return;
    }

    if (config.allowedTenants.length > 0 && !config.allowedTenants.includes(tenantId)) {
      res.status(403).json({ error: 'Tenant not allowed.' });
      return;
    }

    const expectedIssuer = issuerForTenant(tenantId);
    if (decoded.iss !== expectedIssuer) {
      res.status(401).json({ error: 'Invalid issuer.' });
      return;
    }

    req.auth = {
      tenantId,
      userPrincipalName: decoded.preferred_username ?? decoded.upn ?? decoded.email,
      token
    };

    logger.info('auth_success', { user: req.auth.userPrincipalName, tenant: tenantId });
    next();
  });
};
