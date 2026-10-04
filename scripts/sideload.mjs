// Installs (or removes) the add-in manifest so Excel shows Excel Coach.
//   npm run sideload     install for Excel on this Mac
//   npm run unsideload   remove it
// On Windows and Excel on the web, Excel installs from the manifest directly;
// this script prints the steps.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const manifest = path.join(root, 'manifest.xml');
const remove = process.argv.includes('--remove');

if (process.platform === 'darwin') {
  const wef = path.join(os.homedir(), 'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
  const target = path.join(wef, 'excel-coach.manifest.xml');
  if (remove) {
    fs.rmSync(target, { force: true });
    console.log('Removed Excel Coach. Restart Excel to finish.');
  } else {
    fs.mkdirSync(wef, { recursive: true });
    fs.copyFileSync(manifest, target);
    console.log(`Installed Excel Coach for Excel on this Mac:\n  ${target}`);
    console.log('Restart Excel, open a workbook, then Home > Excel Coach.');
    console.log('(If the button is missing: Home > Add-ins, and pick Excel Coach under Developer add-ins.)');
  }
} else {
  console.log('Windows: share a folder that holds manifest.xml, add it under File > Options > Trust Center >');
  console.log('Trust Center Settings > Trusted Add-in Catalogs, restart Excel, then Home > Add-ins > Shared Folder.');
  console.log('Excel on the web: Home > Add-ins > More add-ins > My Add-ins > Upload My Add-in, and pick manifest.xml.');
  console.log('Both need the panel hosted at an HTTPS address that machine can reach.');
}
