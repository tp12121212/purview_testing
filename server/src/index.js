import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import morgan from 'morgan';
import multer from 'multer';
import fs from 'node:fs/promises';
import path from 'node:path';
import { v4 as uuidv4 } from 'uuid';
import { config } from './config.js';
import { logger } from './logger.js';
import { authenticate, isAudienceAllowed, tokenAudiences, verifySupplementalToken } from './auth.js';
import { listSensitiveInformationTypes, runDataClassification, runTextExtraction } from './powershell.js';

const app = express();

app.use(helmet());
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: '1mb' }));
app.use(morgan('combined'));

const upload = multer({
  storage: multer.diskStorage({
    destination: async (req, file, cb) => {
      try {
        await fs.mkdir(config.tempDir, { recursive: true });
        cb(null, config.tempDir);
      } catch (error) {
        cb(error);
      }
    },
    filename: (req, file, cb) => {
      const safeName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
      cb(null, `${uuidv4()}_${safeName}`);
    }
  }),
  limits: { fileSize: config.fileUploadLimitMb * 1024 * 1024 }
});

const validateContentType = (req, res, next) => {
  const { file } = req;
  if (!file) {
    res.status(400).json({ error: 'File is required.' });
    return;
  }

  if (!config.allowedContentTypes.includes(file.mimetype)) {
    res.status(400).json({ error: `Unsupported content type: ${file.mimetype}` });
    return;
  }

  next();
};

const withCleanup = (handler) => async (req, res) => {
  const filePath = req.file?.path;
  try {
    await handler(req, res);
  } finally {
    if (filePath) {
      await fs.unlink(filePath).catch(() => undefined);
    }
  }
};

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.get('/api/sensitive-information-types', authenticate(tokenAudiences.compliance), async (req, res) => {
  logger.info('sits_list_requested', { user: req.auth?.userPrincipalName, tenant: req.auth?.tenantId });
  try {
    const data = await listSensitiveInformationTypes({
      accessToken: req.auth?.token,
      userPrincipalName: req.auth?.userPrincipalName
    });
    res.json({ items: data ?? [] });
  } catch (error) {
    logger.error('sits_list_failed', { error: error.message });
    res.status(500).json({ error: 'Failed to list sensitive information types.' });
  }
});

app.post('/api/extraction', authenticate(tokenAudiences.exchange), upload.single('file'), validateContentType, withCleanup(async (req, res) => {
  logger.info('text_extraction_requested', { user: req.auth?.userPrincipalName, tenant: req.auth?.tenantId });
  try {
    const result = await runTextExtraction({
      filePath: req.file.path,
      accessToken: req.auth?.token,
      userPrincipalName: req.auth?.userPrincipalName
    });
    res.json({ result });
  } catch (error) {
    logger.error('text_extraction_failed', { error: error.message });
    res.status(500).json({ error: 'Text extraction failed.' });
  }
}));

app.post('/api/classification', authenticate(tokenAudiences.classification), upload.single('file'), validateContentType, withCleanup(async (req, res) => {
  logger.info('data_classification_requested', { user: req.auth?.userPrincipalName, tenant: req.auth?.tenantId });

  const { selectedSits, useAllSits } = req.body ?? {};
  const normalizedSelectedSits = typeof selectedSits === 'string'
    ? selectedSits.split(',').map((item) => item.trim()).filter(Boolean)
    : Array.isArray(selectedSits)
      ? selectedSits
      : [];
  const normalizedUseAllSits = typeof useAllSits === 'string' ? useAllSits === 'true' : Boolean(useAllSits);
  const primaryAudience = req.auth?.audience;
  const normalizeHeaderToken = (value) => Array.isArray(value) ? value[0] : value;
  const primaryIsExchange = isAudienceAllowed(primaryAudience, tokenAudiences.exchange);
  const primaryIsCompliance = isAudienceAllowed(primaryAudience, tokenAudiences.compliance);

  const exchangeToken = primaryIsExchange ? req.auth?.token : normalizeHeaderToken(req.headers['x-exchange-token']);
  const complianceToken = primaryIsCompliance ? req.auth?.token : normalizeHeaderToken(req.headers['x-compliance-token']);

  if (!exchangeToken || !complianceToken) {
    res.status(401).json({ error: 'Both Exchange and Compliance tokens are required.' });
    return;
  }

  try {
    if (!primaryIsExchange) {
      await verifySupplementalToken(exchangeToken, {
        expectedAudiences: tokenAudiences.exchange,
        expectedTenantId: req.auth?.tenantId
      });
    }
    if (!primaryIsCompliance) {
      await verifySupplementalToken(complianceToken, {
        expectedAudiences: tokenAudiences.compliance,
        expectedTenantId: req.auth?.tenantId
      });
    }
  } catch (error) {
    logger.warn('classification_token_invalid', { error: error.message });
    res.status(401).json({ error: 'Invalid Exchange or Compliance token.' });
    return;
  }

  try {
    const result = await runDataClassification({
      filePath: req.file.path,
      exchangeAccessToken: exchangeToken,
      complianceAccessToken: complianceToken,
      userPrincipalName: req.auth?.userPrincipalName,
      sensitiveTypes: normalizedSelectedSits,
      useAll: normalizedUseAllSits
    });
    res.json({ result });
  } catch (error) {
    logger.error('data_classification_failed', { error: error.message });
    res.status(500).json({ error: 'Data classification failed.' });
  }
}));

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    res.status(400).json({ error: err.message });
    return;
  }

  logger.error('unhandled_error', { error: err.message });
  res.status(500).json({ error: 'Unexpected server error.' });
});

app.listen(config.port, () => {
  logger.info('server_started', { port: config.port });
});
