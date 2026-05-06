import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.js?url';
import mammoth from 'mammoth/mammoth.browser';
import MsgReader from 'msgreader';
import { createWorker } from 'tesseract.js';
import { Archive } from 'libarchive.js/dist/libarchive.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'tif', 'tiff', 'bmp', 'gif', 'webp']);
const ARCHIVE_EXTENSIONS = new Set(['zip', '7z', 'rar', 'tar', 'tgz', 'gz', 'xz', 'bz2']);
const TEXT_EXTENSIONS = new Set(['txt', 'csv', 'md', 'json', 'eml']);

const MAX_ARCHIVE_DEPTH = 2;
const archiveWorkerUrl = '/libarchive.js/dist/worker-bundle.js';
// Worker bundle must be served from the public directory so libarchive.js can load it.
Archive.init({ workerUrl: archiveWorkerUrl });

let ocrWorkerPromise = null;

const getOcrWorker = () => {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = (async () => {
      const worker = createWorker();
      await worker.load();
      await worker.loadLanguage('eng');
      await worker.initialize('eng');
      return worker;
    })();
  }
  return ocrWorkerPromise;
};

const runOcr = async (target) => {
  const worker = await getOcrWorker();
  const { data } = await worker.recognize(target);
  return data?.text ?? '';
};

const decodeText = (buffer) => new TextDecoder('utf-8').decode(buffer);

const decodeMaybeText = (value) => {
  if (!value) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (value instanceof ArrayBuffer) {
    return decodeText(value);
  }
  if (ArrayBuffer.isView(value)) {
    return decodeText(value.buffer);
  }
  if (value.buffer && value.buffer instanceof ArrayBuffer) {
    return decodeText(value.buffer);
  }
  return String(value);
};

const convertRtfToText = (value) => {
  const rtf = decodeMaybeText(value);
  if (!rtf) {
    return '';
  }
  let text = rtf;
  text = text.replace(/\\par[d]?/gi, '\n');
  text = text.replace(/\\line\b/gi, '\n');
  text = text.replace(/\\'[0-9a-f]{2}/gi, (match) => {
    const hex = match.slice(2);
    return String.fromCharCode(Number.parseInt(hex, 16));
  });
  text = text.replace(/\\u(-?\d+)\??/g, (_, num) => {
    const valueNumber = Number(num);
    const codePoint = valueNumber < 0 ? 65536 + valueNumber : valueNumber;
    if (!Number.isFinite(codePoint)) {
      return '';
    }
    return String.fromCharCode(codePoint);
  });
  text = text.replace(/\\\\/g, '\\');
  text = text.replace(/\\[a-z]+-?\d*/gi, '');
  text = text.replace(/[{}]/g, '');
  text = text.replace(/[\r\n]+/g, '\n');
  return text.replace(/\s+/g, ' ').trim();
};

const stripHtml = (html) => {
  if (!html) {
    return '';
  }
  const container = document.createElement('div');
  container.innerHTML = html;
  return container.textContent ?? '';
};

const getExtension = (name = '') => {
  const parts = name.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
};

const flattenArchiveEntries = (node, prefix = '') => {
  const entries = [];
  if (!node || typeof node !== 'object') {
    return entries;
  }
  Object.entries(node).forEach(([key, value]) => {
    const entryName = prefix ? `${prefix}/${key}` : key;
    if (value instanceof File) {
      entries.push({ file: value, name: entryName });
    } else if (value && typeof value === 'object') {
      entries.push(...flattenArchiveEntries(value, entryName));
    }
  });
  return entries;
};

const renderPageToCanvas = async (page) => {
  const viewport = page.getViewport({ scale: 2 });
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas;
};

const extractTextFromPdf = async (buffer) => {
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const textByPage = [];
  let totalLength = 0;
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(' ');
    textByPage.push(pageText);
    totalLength += pageText.trim().length;
  }
  const extracted = textByPage.join('\n');
  if (totalLength > 10) {
    return extracted;
  }
  const ocrPages = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const canvas = await renderPageToCanvas(page);
    const ocrText = await runOcr(canvas);
    if (ocrText.trim()) {
      ocrPages.push(ocrText);
    }
  }
  return ocrPages.join('\n\n');
};

const extractTextFromDocx = async (buffer) => {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value ?? '';
};

const extractTextFromMsg = (buffer) => {
  try {
    const msgReader = new MsgReader(buffer);
    const message = msgReader.getFileData();
    const html = stripHtml(decodeMaybeText(message.bodyHTML));
    if (html) {
      return html;
    }
    const rtf = convertRtfToText(message.bodyRTF ?? message.bodyRtf ?? message.body);
    if (rtf) {
      return rtf;
    }
    const plainBody = decodeMaybeText(message.body);
    if (plainBody) {
      return plainBody;
    }
    return message.subject ?? '';
  } catch {
    return '';
  }
};

const extractTextFromImage = async (buffer) => {
  const blob = new Blob([buffer]);
  return (await runOcr(blob)).trim();
};

async function extractTextFromArchive({ buffer, name, depth }) {
  if (depth >= MAX_ARCHIVE_DEPTH) {
    throw new Error('Archive nesting depth exceeded.');
  }
  const archiveFile = new File([buffer], name || 'archive', {
    type: 'application/octet-stream'
  });
  const archive = await Archive.open(archiveFile);
  try {
    const extracted = await archive.extractFiles();
    const entries = flattenArchiveEntries(extracted);
    if (!entries.length) {
      throw new Error('Archive contained no readable entries.');
    }
    const texts = [];
    for (const entry of entries) {
      const entryBuffer = await entry.file.arrayBuffer();
      const entryText = await extractTextFromBuffer({
        buffer: entryBuffer,
        name: entry.name,
        depth: depth + 1
      });
      if (entryText) {
        texts.push(`=== ${entry.name} ===\n${entryText}`);
      }
    }
    if (!texts.length) {
      throw new Error('Archive contained no supported files.');
    }
    return texts.join('\n\n');
  } finally {
    await archive.close();
  }
}

export const extractTextFromFile = async (file) => {
  if (!file) {
    throw new Error('No file provided.');
  }
  const buffer = await file.arrayBuffer();
  return extractTextFromBuffer({ buffer, name: file.name, depth: 0 });
};

export const extractTextFromBuffer = async ({ buffer, name = '', depth = 0 }) => {
  const extension = getExtension(name);
  if (!extension) {
    throw new Error('File must include an extension.');
  }
  if (ARCHIVE_EXTENSIONS.has(extension)) {
    return extractTextFromArchive({ buffer, name, depth });
  }
  if (extension === 'pdf') {
    return extractTextFromPdf(buffer);
  }
  if (extension === 'docx') {
    return extractTextFromDocx(buffer);
  }
  if (extension === 'msg') {
    return extractTextFromMsg(buffer);
  }
  if (IMAGE_EXTENSIONS.has(extension)) {
    return extractTextFromImage(buffer);
  }
  if (TEXT_EXTENSIONS.has(extension)) {
    return decodeText(buffer);
  }
  throw new Error(`Unsupported file type: .${extension}.`);
};

export const extractTextFromHtml = (html) => stripHtml(html);
