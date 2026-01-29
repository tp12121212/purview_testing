import fs from 'node:fs';
import path from 'node:path';

const configPath = path.join(process.cwd(), 'web', 'public', 'runtime-config.json');

const requiredKeys = [
  'clientId',
  'authorityHost',
  'authorityTenant',
  'redirectUri',
  'graphBaseUrl',
  'graphApiVersion',
  'infoProtectionApiVersion',
  'loginScopes'
];

const run = () => {
  if (!fs.existsSync(configPath)) {
    console.error(`Missing runtime config at ${configPath}`);
    process.exit(1);
  }

  const raw = fs.readFileSync(configPath, 'utf8');
  let config;
  try {
    config = JSON.parse(raw);
  } catch (error) {
    console.error('runtime-config.json is not valid JSON.');
    process.exit(1);
  }

  const missing = requiredKeys.filter((key) => !(key in config));
  if (missing.length) {
    console.error(`runtime-config.json missing keys: ${missing.join(', ')}`);
    process.exit(1);
  }

  if (!config.clientId) {
    console.warn('Warning: clientId is empty. Sign-in will be disabled until configured.');
  }

  console.log('runtime-config.json OK');
};

run();
