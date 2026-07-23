import { mkdir, writeFile, cp } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { chromium } from 'playwright';
import { content } from './content.mjs';
import { renderHTML } from './templates.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.join(root, 'dist');
const outDir = path.join(root, 'output');

await mkdir(distDir, { recursive: true });
await mkdir(outDir, { recursive: true });
await cp(path.join(root, 'style.css'), path.join(distDir, 'style.css'));
await cp(path.join(root, 'assets'), path.join(distDir, 'assets'), { recursive: true });

const html = renderHTML(content);
const htmlPath = path.join(distDir, 'index.html');
await writeFile(htmlPath, html, 'utf-8');

const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto('file://' + htmlPath, { waitUntil: 'networkidle' });
await page.pdf({
  path: path.join(outDir, 'fcs-zytig.pdf'),
  format: 'A4',
  printBackground: true,
  margin: { top: 0, bottom: 0, left: 0, right: 0 },
});
await browser.close();

console.log('PDF erstellt: output/fcs-zytig.pdf');
