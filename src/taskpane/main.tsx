import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { detectPlatform } from '../engine/platform';
import { App, type Env } from './App';

let started = false;

function start(host: Office.HostType | null, officePlatform: Office.PlatformType | null) {
  if (started) return;
  started = true;
  const params = new URLSearchParams(location.search);
  const officeLoaded = typeof Office !== 'undefined';
  const inExcel = officeLoaded && host != null && host === Office.HostType.Excel;
  const env: Env = {
    inExcel,
    platform: detectPlatform(officePlatform == null ? null : String(officePlatform)),
    mock: params.has('mock'),
    excelApi113: inExcel && Office.context.requirements.isSetSupported('ExcelApi', '1.13'),
  };
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App env={env} />
    </StrictMode>,
  );
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => start(info.host, info.platform));
}
// If office.js can't load (offline, or opened outside Office), render anyway.
setTimeout(() => start(null, null), 3000);
