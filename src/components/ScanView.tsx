import { Badge, Body1, Button, Caption1, Link, MessageBar, MessageBarBody, Spinner, Text, Tooltip, makeStyles, shorthands, tokens } from '@fluentui/react-components';
import { ArrowLeft16Regular, ArrowRight16Regular, LockClosed16Regular, ScanText20Regular, Wrench16Regular } from '@fluentui/react-icons';
import { useEffect, useMemo, useRef, useState } from 'react';
import { FIXABLE, planFix } from '../engine/fix';
import { scanWorkbook, type Finding, type SheetFormulas } from '../engine/scan';
import { friendlyError } from '../excel/errors';
import { SHEET_PREFIX } from '../excel/host';
import { getExercise } from '../exercises';
import { useCoach } from '../taskpane/context';
import { SectionLabel } from './bits';

const useStyles = makeStyles({
  page: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL} ${tokens.spacingVerticalXXL}` },
  back: { alignSelf: 'flex-start', marginLeft: `calc(-1 * ${tokens.spacingHorizontalS})` },
  head: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS },
  muted: { color: tokens.colorNeutralForeground3 },
  privacy: { display: 'grid', gridTemplateColumns: '16px minmax(0, 1fr)', gap: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground2, alignItems: 'start' },
  summary: { display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  list: { listStyle: 'none', margin: 0, padding: 0, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  cardHead: { display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  fix: { color: tokens.colorNeutralForeground1 },
  cells: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS, marginTop: tokens.spacingVerticalXXS },
  cell: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    lineHeight: tokens.lineHeightBase200,
    padding: '1px 6px',
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    background: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground2,
    cursor: 'pointer',
    ':hover': { ...shorthands.borderColor(tokens.colorBrandStroke1), color: tokens.colorBrandForeground1 },
    ':focus-visible': { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: '1px' },
  },
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalM,
    rowGap: tokens.spacingVerticalXS,
    marginTop: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke3}`,
  },
  practice: { display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: tokens.fontSizeBase200, textAlign: 'left' },
  empty: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, padding: tokens.spacingHorizontalM, borderRadius: tokens.borderRadiusLarge, backgroundColor: tokens.colorNeutralBackground2 },
});

const SEVERITY: Record<Finding['severity'], { label: string; color: 'danger' | 'warning' | 'informative' }> = {
  high: { label: 'Fix first', color: 'danger' },
  medium: { label: 'Worth fixing', color: 'warning' },
  low: { label: 'Tidy up', color: 'informative' },
};

interface ScanResult {
  workbook: string;
  /** The sheets as scanned, kept so a finding can become a guided fix. */
  sheets: SheetFormulas[];
  truncated: boolean;
  findings: Finding[];
}

/**
 * The last scan in this panel, shown again when the learner comes back from a guided fix. By then it
 * can be out of date: a finished fix ends with pasting the new formulas onto the original sheet. So
 * the view scans again as it opens, and shows this list until the new one is ready.
 */
let lastResult: ScanResult | undefined;

/** The first sheet, in the order of the finding's examples, where it has a guided fix. */
function fixableSheet(f: Finding, sheets: SheetFormulas[]): SheetFormulas | undefined {
  if (!FIXABLE.has(f.id)) return undefined;
  const names = [...new Set(f.cells.map((c) => c.sheet))];
  for (const name of names) {
    const sheet = sheets.find((x) => x.sheet === name);
    if (sheet && planFix(f, sheet)) return sheet;
  }
  return undefined;
}

export function ScanView() {
  const s = useStyles();
  const { host, hostNotice, goHome, openExercise, openFix } = useCoach();
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const [result, setResult] = useState<ScanResult | undefined>(() => lastResult);
  const refreshed = useRef(false);

  const scan = async () => {
    if (!host) return setError(hostNotice);
    setBusy(true);
    setError(undefined);
    try {
      const wb = await host.readWorkbook();
      const next = { workbook: wb.name, sheets: wb.sheets, truncated: wb.truncated, findings: scanWorkbook(wb.sheets, { skipPrefix: SHEET_PREFIX }) };
      lastResult = next;
      setResult(next);
    } catch (e) {
      setError(friendlyError(e));
    } finally {
      setBusy(false);
    }
  };

  // Scanning is local and quick, so a list from earlier is refreshed on opening, once.
  useEffect(() => {
    if (refreshed.current || !lastResult || !host) return;
    refreshed.current = true;
    void scan();
    // Only on opening: `scan` is a new function every render.
  }, []);

  const fixable = useMemo(() => new Map(result?.findings.map((f) => [f.id, fixableSheet(f, result.sheets)] as const)), [result]);
  const goTo = (sheet: string, address: string) => host?.select(sheet, address).catch((e) => setError(friendlyError(e)));
  const total = result?.findings.reduce((a, f) => a + f.count, 0) ?? 0;
  const sheetCount = result?.sheets.length ?? 0;

  return (
    <div className={s.page}>
      <Button className={s.back} appearance="transparent" size="small" icon={<ArrowLeft16Regular />} onClick={goHome}>
        All skills
      </Button>

      <div className={s.head}>
        <Text as="h1" size={500} weight="semibold" style={{ margin: 0 }}>
          My Work
        </Text>
        <Body1 className={s.muted}>Find the spots in your own workbook where an advanced skill would save time or prevent mistakes.</Body1>
      </div>

      <div className={s.privacy}>
        <LockClosed16Regular aria-hidden="true" style={{ marginTop: 2 }} />
        <Caption1>The scan reads formulas in the open workbook, on this computer. Nothing is uploaded, and nothing in the workbook changes.</Caption1>
      </div>

      {error && (
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <Button appearance={result ? 'secondary' : 'primary'} size="large" icon={busy ? <Spinner size="tiny" /> : <ScanText20Regular />} disabled={busy} onClick={() => void scan()}>
        {busy ? 'Scanning…' : result ? 'Scan again' : 'Scan this workbook'}
      </Button>

      {result && (
        <section aria-label="Results" aria-busy={busy} style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
          <div className={s.summary}>
            <SectionLabel>{result.findings.length ? `${result.findings.length} kinds of upgrade spot` : 'Results'}</SectionLabel>
            <Caption1 className={s.muted}>
              {result.workbook} · {sheetCount} {sheetCount === 1 ? 'sheet' : 'sheets'}
            </Caption1>
          </div>

          {sheetCount === 0 ? (
            <div className={s.empty}>
              <Text weight="semibold">Nothing to scan here</Text>
              <Caption1>This workbook only has practice sheets. Open one of your own workbooks, open Excel Coach there, and scan it.</Caption1>
            </div>
          ) : result.findings.length === 0 ? (
            <div className={s.empty}>
              <Text weight="semibold">No upgrade spots found</Text>
              <Caption1>The scan looks for fragile lookups, broken fills, values typed over formulas, hard-coded numbers and ranges that should be Tables.</Caption1>
            </div>
          ) : (
            <ul className={s.list}>
              {result.findings.map((f) => {
                const ex = f.exerciseId ? getExercise(f.exerciseId) : undefined;
                const fixSheet = fixable.get(f.id);
                return (
                  <li key={f.id} className={s.card}>
                    <div className={s.cardHead}>
                      <Text weight="semibold">{f.title}</Text>
                      <Badge appearance="tint" color={SEVERITY[f.severity].color} size="small" style={{ flexShrink: 0 }}>
                        {SEVERITY[f.severity].label}
                      </Badge>
                    </div>
                    <Caption1 className={s.muted}>{f.why}</Caption1>
                    <Caption1 className={s.fix}>{f.fix}</Caption1>
                    <div className={s.cells}>
                      {f.cells.map((c) => (
                        <button key={`${c.sheet}!${c.address}`} className={s.cell} onClick={() => goTo(c.sheet, c.address)} title={`Go to ${c.sheet}!${c.address}`}>
                          {sheetCount > 1 ? `${c.sheet}!` : ''}
                          {c.address}
                        </button>
                      ))}
                      {f.count > f.cells.length && <Caption1 className={s.muted}>+{f.count - f.cells.length} more</Caption1>}
                    </div>
                    {(fixSheet || ex) && (
                      <div className={s.actions}>
                        {/* Off while a scan runs: the plan would come from the sheets as they were. */}
                        {fixSheet && (
                          <Tooltip content={`Copies ${fixSheet.sheet} and walks you through the fix there. Your sheet isn’t changed.`} relationship="description">
                            <Button size="small" icon={<Wrench16Regular />} disabled={busy} onClick={() => openFix({ finding: f, sheet: fixSheet, workbook: result.workbook })}>
                              Fix it on a copy
                            </Button>
                          </Tooltip>
                        )}
                        {ex && (
                          <Link
                            as="button"
                            className={s.practice}
                            onClick={() => openExercise(ex.id)}
                            aria-label={fixSheet ? `Practice the fix first: ${ex.title}` : undefined}
                          >
                            {fixSheet ? 'Practice the fix first' : `Practice the fix: ${ex.title}`} <ArrowRight16Regular />
                          </Link>
                        )}
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>
          )}
          {result.truncated && <Caption1 className={s.muted}>Very large sheets were scanned in part, starting from the top.</Caption1>}
          {total > 0 && <Caption1 className={s.muted}>{total} cells in all. Select a cell reference to jump to it.</Caption1>}
        </section>
      )}
    </div>
  );
}
