import type { Platform } from './types';

const isMacOS = () => typeof navigator !== 'undefined' && /Mac/i.test(navigator.platform || navigator.userAgent);

/** Maps Office's platform string to ours; falls back to the browser's OS. */
export function detectPlatform(officePlatform?: string | null): Platform {
  switch (officePlatform) {
    case 'Mac':
      return 'mac';
    case 'PC':
      return 'windows';
    case 'OfficeOnline':
      return 'web';
    default:
      return isMacOS() ? 'mac' : 'windows';
  }
}

interface Keys {
  tableKey: string;
  absKey: string;
  nameManager: string;
  enter: string;
  /** Where macro recording starts. */
  recordMacro: string;
  /** Where a saved macro is run from. */
  runMacro: string;
  /** How to open the VBA editor. */
  vbaEditor: string;
  /** Where the Developer tab is turned on. */
  developerTab: string;
  /** The jump-to-edge shortcut that End(xlUp) mimics. */
  jumpUp: string;
  /** How the Personal Macro Workbook's project appears in the VBA editor. */
  personalProject: string;
  /** Fill Right: copies a formula across without sliding Table column names, unlike dragging. */
  fillRight: string;
}

const NO_VBA = 'Excel on a Mac or Windows PC (Excel for the web can’t run VBA)';

const KEYS: Record<Platform, Keys> = {
  mac: {
    tableKey: '⌘T',
    absKey: '⌘T',
    nameManager: 'Formulas › Define Name',
    enter: 'Return',
    recordMacro: 'Tools › Macro › Record New Macro',
    runMacro: 'Tools › Macro › Macros (⌥F8)',
    vbaEditor: 'Tools › Macro › Visual Basic Editor (⌥F11)',
    developerTab: 'Excel › Settings › Ribbon & Toolbar',
    jumpUp: '⌘↑',
    personalProject: 'your Personal Macro Workbook’s project',
    fillRight: '⌘R',
  },
  windows: {
    tableKey: 'Ctrl+T',
    absKey: 'F4',
    nameManager: 'Formulas › Name Manager',
    enter: 'Enter',
    recordMacro: 'View › Macros › Record Macro',
    runMacro: 'View › Macros › View Macros (Alt+F8)',
    vbaEditor: 'the Visual Basic Editor (Alt+F11)',
    developerTab: 'File › Options › Customize Ribbon',
    jumpUp: 'Ctrl+↑',
    personalProject: 'VBAProject (PERSONAL.XLSB)',
    fillRight: 'Ctrl+R',
  },
  web: {
    tableKey: 'Ctrl+T',
    absKey: 'F4',
    nameManager: 'Formulas › Name Manager',
    enter: 'Enter',
    recordMacro: NO_VBA,
    runMacro: NO_VBA,
    vbaEditor: NO_VBA,
    developerTab: NO_VBA,
    jumpUp: 'Ctrl+↑',
    personalProject: NO_VBA,
    fillRight: 'Ctrl+R',
  },
};

/**
 * Where a Power Query starts from a Table. Excel for Mac has no From Table/Range, so a Blank
 * query reads the Table with Excel.CurrentWorkbook() instead.
 */
function fromTable(table: string, platform: Platform): string {
  return platform === 'mac'
    ? `Data › Get Data (Power Query) › Blank query, replace the text with \`Excel.CurrentWorkbook(){[Name="${table}"]}[Content]\` and select Next`
    : `Data › From Table/Range (with a cell in \`${table}\` selected)`;
}

/**
 * Naming a helper query so Merge and Append list it by its Table's name. From Table/Range names the
 * query after the Table; a Blank query on Mac starts as "Query (n)".
 */
function nameQuery(table: string, platform: Platform): string {
  return platform === 'mac' ? `name it ${table} in Query Settings › Name (a Blank query starts as Query), ` : '';
}

const MAC_PQ_NOTE =
  'Excel for Mac has no From Table/Range: start each query with Data › Get Data (Power Query) › Blank query and the one-line starter in the first hint. After that the editor works the same.';

/**
 * Replaces platform placeholders: {tableKey}, {absKey}, {nameManager}, {enter}, {recordMacro},
 * {runMacro}, {vbaEditor}, {developerTab}, {jumpUp}, {personalProject}, {fillRight}, {fromTable:TableName}, {nameQuery:TableName} and {macPqNote} (the last two
 * are empty outside Excel for Mac).
 */
export function localize(text: string, platform: Platform): string {
  const keys = KEYS[platform];
  return text
    .replace(/\{fromTable:([^}]+)\}/g, (_, table: string) => fromTable(table, platform))
    .replace(/\{nameQuery:([^}]+)\}/g, (_, table: string) => nameQuery(table, platform))
    .replace(/\s?\{macPqNote\}/g, platform === 'mac' ? ` ${MAC_PQ_NOTE}` : '')
    .replace(/\{(tableKey|absKey|nameManager|enter|recordMacro|runMacro|vbaEditor|developerTab|jumpUp|personalProject|fillRight)\}/g, (_, k: keyof Keys) => keys[k]);
}
