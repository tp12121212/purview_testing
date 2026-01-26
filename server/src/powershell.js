import { spawn } from 'node:child_process';
import path from 'node:path';
import { config } from './config.js';
import { logger } from './logger.js';

const runScript = ({ scriptName, args = [], inputJson = null }) => new Promise((resolve, reject) => {
  const scriptPath = path.join(config.powershellScriptsDir, scriptName);
  const commandArgs = ['-NoLogo', '-NoProfile', '-File', scriptPath, ...args];

  const child = spawn(config.powershellPath, commandArgs, { stdio: ['pipe', 'pipe', 'pipe'] });

  let stdout = '';
  let stderr = '';

  child.stdout.on('data', (data) => {
    stdout += data.toString();
  });

  child.stderr.on('data', (data) => {
    stderr += data.toString();
  });

  child.on('error', (error) => {
    reject(error);
  });

  child.on('close', (code) => {
    if (code !== 0) {
      logger.error('powershell_failed', { scriptName, code, stderr });
      reject(new Error(stderr || `PowerShell exited with code ${code}`));
      return;
    }

    resolve({ stdout, stderr });
  });

  if (inputJson) {
    child.stdin.write(JSON.stringify(inputJson));
  }
  child.stdin.end();
});

const parseJson = (output) => {
  const trimmed = output.trim();
  if (!trimmed) {
    return null;
  }
  return JSON.parse(trimmed);
};

export const runTextExtraction = async ({ filePath, accessToken, userPrincipalName }) => {
  const args = ['-FilePath', filePath];
  if (accessToken) {
    args.push('-AccessToken', accessToken);
  }
  if (userPrincipalName) {
    args.push('-UserPrincipalName', userPrincipalName);
  }

  const { stdout } = await runScript({ scriptName: 'text-extraction.ps1', args });
  return parseJson(stdout);
};

export const runDataClassification = async ({ filePath, accessToken, userPrincipalName, sensitiveTypes, useAll }) => {
  const args = ['-FilePath', filePath];
  if (accessToken) {
    args.push('-AccessToken', accessToken);
  }
  if (userPrincipalName) {
    args.push('-UserPrincipalName', userPrincipalName);
  }
  if (useAll) {
    args.push('-AllSensitiveInformationTypes');
  }
  if (sensitiveTypes?.length) {
    args.push('-SensitiveInformationTypes', sensitiveTypes.join(','));
  }

  const { stdout } = await runScript({ scriptName: 'data-classification.ps1', args });
  return parseJson(stdout);
};

export const listSensitiveInformationTypes = async ({ accessToken, userPrincipalName }) => {
  const args = [];
  if (accessToken) {
    args.push('-AccessToken', accessToken);
  }
  if (userPrincipalName) {
    args.push('-UserPrincipalName', userPrincipalName);
  }
  const { stdout } = await runScript({ scriptName: 'list-sits.ps1', args });
  return parseJson(stdout);
};
