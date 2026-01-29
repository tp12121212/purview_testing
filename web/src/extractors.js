import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import mammoth from 'mammoth/mammoth.browser';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const decodeText = (buffer) => new TextDecoder('utf-8').decode(buffer);

const stripHtml = (html) => {
  if (!html) {
    return '';
  }
  const container = document.createElement('div');
  container.innerHTML = html;
  return container.textContent ?? '';
};

const extractTextFromPdf = async (buffer) => {
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(' ');
    pages.push(pageText);
  }
  return pages.join('\n');
};

const extractTextFromDocx = async (buffer) => {
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value ?? '';
};

const getExtension = (name) => {
  if (!name) {
    return '';
  }
  const parts = name.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
};

export const extractTextFromFile = async (file) => {
  if (!file) {
    throw new Error('No file provided.');
  }
  const extension = getExtension(file.name);
  const buffer = await file.arrayBuffer();

  if (extension === 'pdf') {
    return extractTextFromPdf(buffer);
  }
  if (extension === 'docx') {
    return extractTextFromDocx(buffer);
  }
  if (extension === 'txt' || extension === 'csv' || extension === 'md' || extension === 'json') {
    return decodeText(buffer);
  }
  if (extension === 'eml') {
    return decodeText(buffer);
  }

  throw new Error(`Unsupported file type: .${extension}.`);
};

export const extractTextFromBuffer = async ({ buffer, name }) => {
  const extension = getExtension(name);
  if (extension === 'pdf') {
    return extractTextFromPdf(buffer);
  }
  if (extension === 'docx') {
    return extractTextFromDocx(buffer);
  }
  return decodeText(buffer);
};

export const extractTextFromHtml = (html) => stripHtml(html);
