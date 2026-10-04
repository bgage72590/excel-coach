import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Body1, Caption1, makeStyles, tokens } from '@fluentui/react-components';
import { BookOpen20Regular, Lightbulb16Regular } from '@fluentui/react-icons';
import type { Concept } from '../engine/types';
import { RichText } from './bits';

const useStyles = makeStyles({
  wrap: {
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
  },
  panel: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, padding: `0 ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}` },
  syntax: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    overflowWrap: 'anywhere',
    margin: 0,
    whiteSpace: 'pre-wrap',
  },
  example: { color: tokens.colorNeutralForeground2 },
  tip: { display: 'grid', gridTemplateColumns: '16px minmax(0, 1fr)', gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground2 },
  tipIcon: { color: tokens.colorPaletteMarigoldForeground1, marginTop: '1px' },
});

export function ConceptPanel({ concept, defaultOpen }: { concept: Concept; defaultOpen: boolean }) {
  const s = useStyles();
  return (
    <div className={s.wrap}>
      <Accordion collapsible defaultOpenItems={defaultOpen ? ['concept'] : []}>
        <AccordionItem value="concept">
          <AccordionHeader icon={<BookOpen20Regular />} expandIconPosition="end" size="medium">
            How it works
          </AccordionHeader>
          <AccordionPanel>
            <div className={s.panel}>
              <Body1>{concept.summary}</Body1>
              <pre className={s.syntax}>{concept.syntax}</pre>
              {concept.example && (
                <Caption1 className={s.example}>
                  <RichText text={concept.example} />
                </Caption1>
              )}
              {concept.tip && (
                <div className={s.tip}>
                  <Lightbulb16Regular className={s.tipIcon} aria-hidden="true" />
                  <Caption1>{concept.tip}</Caption1>
                </div>
              )}
            </div>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
}
