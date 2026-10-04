import { describe, expect, it } from 'vitest';
import { eomonth, serial } from '../src/engine/data';
import { Rng } from '../src/engine/rng';
import {
  DATES,
  HOLIDAYS,
  MIN_HOLIDAY_LINES,
  REPORT_DATES,
  TERMS,
  daysOverdue,
  dueDate,
  eomonthTerms,
  isWeekend,
  networkdays,
  networkdaysActual,
  workday,
  workdayReceipt,
} from '../src/exercises/dates';
import { gradeRules } from '../src/engine/grade';
import { SEEDS, exerciseSuite } from './helpers/suite';

exerciseSuite(DATES);

const d = serial;

describe('isWeekend', () => {
  it('knows Saturday and Sunday from the serial', () => {
    expect(isWeekend(d(2026, 10, 3))).toBe(true); // Saturday
    expect(isWeekend(d(2026, 10, 4))).toBe(true); // Sunday
    expect(isWeekend(d(2026, 10, 5))).toBe(false); // Monday
    expect(isWeekend(d(2026, 10, 9))).toBe(false); // Friday
  });

  it('every listed holiday is an observed weekday', () => {
    for (const h of HOLIDAYS) expect(isWeekend(h.date), h.name).toBe(false);
  });
});

describe('workday matches Excel WORKDAY', () => {
  const thanksgiving = [d(2026, 11, 26), d(2026, 11, 27)];
  const yearEnd = [d(2026, 12, 24), d(2026, 12, 25), d(2027, 1, 1)];

  it.each([
    ['Friday + 5 across Thanksgiving', d(2026, 11, 20), 5, thanksgiving, d(2026, 12, 1)],
    ['Friday + 1 is Monday', d(2026, 10, 2), 1, [], d(2026, 10, 5)],
    ['Saturday start: + 1 is Monday', d(2026, 10, 3), 1, [], d(2026, 10, 5)],
    ['Sunday start: + 5 is Friday', d(2026, 10, 4), 5, [], d(2026, 10, 9)],
    ['Saturday start before a Monday holiday', d(2026, 9, 5), 1, [d(2026, 9, 7)], d(2026, 9, 8)],
    ['holiday start is not counted', d(2026, 12, 24), 1, yearEnd, d(2026, 12, 28)],
    ['Wednesday + 3 over Christmas', d(2026, 12, 23), 3, yearEnd, d(2026, 12, 30)],
    ['Wednesday + 6 over Christmas and New Year', d(2026, 12, 23), 6, yearEnd, d(2027, 1, 5)],
    ['a weekend holiday is not skipped twice', d(2026, 7, 2), 2, [d(2026, 7, 3), d(2026, 7, 4)], d(2026, 7, 7)],
    ['duplicate holidays count once', d(2026, 11, 20), 5, [...thanksgiving, ...thanksgiving], d(2026, 12, 1)],
    ['0 days returns the start', d(2026, 10, 5), 0, [], d(2026, 10, 5)],
    ['10 days is two weeks', d(2026, 10, 5), 10, [], d(2026, 10, 19)],
    ['-1 from Monday is Friday', d(2026, 10, 5), -1, [], d(2026, 10, 2)],
    ['-1 back over a Monday holiday', d(2026, 9, 8), -1, [d(2026, 9, 7)], d(2026, 9, 4)],
  ])('%s', (_name, start, days, holidays, want) => {
    expect(workday(start, days, holidays)).toBe(want);
  });
});

describe('networkdays matches Excel NETWORKDAYS', () => {
  it.each([
    ['Monday to Friday', d(2026, 10, 5), d(2026, 10, 9), [], 5],
    ['same business day counts 1', d(2026, 10, 5), d(2026, 10, 5), [], 1],
    ['a weekend counts 0', d(2026, 10, 3), d(2026, 10, 4), [], 0],
    ['Friday to Monday counts both ends', d(2026, 10, 2), d(2026, 10, 5), [], 2],
    ['across Thanksgiving', d(2026, 11, 20), d(2026, 12, 1), [d(2026, 11, 26), d(2026, 11, 27)], 6],
    ['end before start is negative', d(2026, 10, 9), d(2026, 10, 5), [], -5],
    ['duplicate and weekend holidays count once', d(2026, 7, 1), d(2026, 7, 7), [d(2026, 7, 3), d(2026, 7, 3), d(2026, 7, 4)], 4],
    ['Saturday to Saturday over Labor Day', d(2026, 9, 5), d(2026, 9, 12), [d(2026, 9, 7)], 4],
    ['a holiday start is not counted', d(2026, 9, 7), d(2026, 9, 9), [d(2026, 9, 7)], 2],
  ])('%s', (_name, start, end, holidays, want) => {
    expect(networkdays(start, end, holidays)).toBe(want);
  });

  it('agrees with workday: n business days on from a business day spans n + 1', () => {
    const holidays = HOLIDAYS.map((h) => h.date);
    const closed = new Set(holidays);
    const rng = new Rng(2026);
    for (let i = 0; i < 400; i++) {
      const start = rng.int(d(2026, 1, 1), d(2026, 12, 31));
      const n = rng.int(0, 40);
      const end = workday(start, n, holidays);
      if (n > 0) {
        expect(isWeekend(end) || closed.has(end)).toBe(false);
        expect(networkdays(start + 1, end, holidays)).toBe(n);
      }
      if (!isWeekend(start) && !closed.has(start)) expect(networkdays(start, end, holidays)).toBe(n + 1);
    }
  });
});

describe('end-of-month terms', () => {
  it('adds the terms to the end of the issue month', () => {
    expect(dueDate(d(2026, 7, 14), 30)).toBe(d(2026, 8, 30));
    expect(dueDate(d(2026, 7, 31), 30)).toBe(d(2026, 8, 30));
    expect(dueDate(d(2026, 6, 3), 60)).toBe(d(2026, 8, 29));
    expect(dueDate(d(2026, 9, 30), 15)).toBe(d(2026, 10, 15));
  });

  it('counts days overdue, and 0 when not overdue or due on the report date', () => {
    expect(daysOverdue(d(2026, 8, 30), d(2026, 9, 30))).toBe(31);
    expect(daysOverdue(d(2026, 10, 15), d(2026, 9, 30))).toBe(0);
    expect(daysOverdue(d(2026, 9, 30), d(2026, 9, 30))).toBe(0);
  });

  it.each(SEEDS)('seed %i: invoices predate every report date, and every terms/report pair mixes overdue and current', (seed) => {
    const data = eomonthTerms.make(new Rng(seed));
    const issued = data.invoices.map((i) => i.issued);
    expect(Math.max(...issued)).toBeLessThanOrEqual(Math.min(...REPORT_DATES));
    expect(issued.some((s) => eomonth(s) === s)).toBe(true);
    for (const terms of TERMS) {
      for (const report of REPORT_DATES) {
        const overdue = eomonthTerms.expected({ ...data, terms, report }).map((r) => r[1] as number);
        expect(overdue.filter((n) => n > 0).length, `terms ${terms}, report ${report}`).toBeGreaterThanOrEqual(3);
        expect(overdue.filter((n) => n === 0).length, `terms ${terms}, report ${report}`).toBeGreaterThanOrEqual(3);
      }
    }
  });
});

describe('practice data', () => {
  const noHolidays = <T extends { holidays: unknown[] }>(data: T): T => ({ ...data, holidays: [] });

  it.each(SEEDS)('seed %i: enough PO lines run through a holiday', (seed) => {
    const data = workdayReceipt.make(new Rng(seed));
    const withH = workdayReceipt.expected(data);
    const without = workdayReceipt.expected(noHolidays(data));
    expect(withH.filter((r, i) => r[0] !== without[i][0]).length).toBeGreaterThanOrEqual(MIN_HOLIDAY_LINES);
  });

  it.each(SEEDS)('seed %i: enough shipments run through a holiday', (seed) => {
    const data = networkdaysActual.make(new Rng(seed));
    const withH = networkdaysActual.expected(data);
    const without = networkdaysActual.expected(noHolidays(data));
    expect(withH.filter((r, i) => r[0] !== without[i][0]).length).toBeGreaterThanOrEqual(MIN_HOLIDAY_LINES);
    for (const s of data.ships) expect(s.delivered).toBeGreaterThan(s.ordered);
  });

  it.each(SEEDS)('seed %i: every variant changes the answer key, not only the inputs', (seed) => {
    for (const ex of DATES) {
      const base = ex.make(new Rng(seed));
      const baseExpected = JSON.stringify(ex.expected(base));
      ex.variants.forEach((v, i) => {
        for (let k = 0; k < 5; k++) {
          const changed = v.apply(base, new Rng(seed + 7919 * (i + 1) + k));
          expect(JSON.stringify(ex.expected(changed)), `${ex.id}: ${v.label}`).not.toBe(baseExpected);
        }
      });
    }
  });

  it('the added holiday is appended as one new row on a business day', () => {
    const data = workdayReceipt.make(new Rng(42));
    const changed = workdayReceipt.variants[0].apply(data, new Rng(5));
    expect(changed.holidays.length).toBe(data.holidays.length + 1);
    expect(changed.holidays.slice(0, -1)).toEqual(data.holidays);
    const added = changed.holidays[changed.holidays.length - 1];
    expect(isWeekend(added.date)).toBe(false);
    expect(data.holidays.some((h) => h.date === added.date)).toBe(false);
  });
});

describe('rules accept good formulas and catch shortcuts', () => {
  const statuses = (ex: { rules?: Parameters<typeof gradeRules>[0] }, formulas: string[]) =>
    gradeRules(ex.rules, formulas).map((i) => `${i.label}: ${i.status}`);
  const allPass = (ex: { rules?: Parameters<typeof gradeRules>[0] }, formulas: string[]) =>
    gradeRules(ex.rules, formulas).every((i) => i.status === 'pass');

  it.each([
    ['=WORKDAY(F2,G2,Holidays[Date])'],
    ['=WORKDAY( F2, G2, Holidays[Date] )'],
    ['=WORKDAY.INTL(F2,G2,1,Holidays[Date])'],
    ['=WORKDAY(F2,G2,Holidays[[#Data],[Date]])'],
  ])('workday accepts %s', (f) => expect(allPass(workdayReceipt, [f]), statuses(workdayReceipt, [f]).join('; ')).toBe(true));

  it('workday flags a fixed holiday range and typed lead days', () => {
    expect(allPass(workdayReceipt, ['=WORKDAY(F2,G2,$B$2:$B$15)'])).toBe(false);
    expect(allPass(workdayReceipt, ['=WORKDAY(F2,10,Holidays[Date])'])).toBe(false);
  });

  it.each([['=NETWORKDAYS(F2,G2,Holidays[Date])'], ['=NETWORKDAYS.INTL(F2,G2,1,Holidays[Date])']])('networkdays accepts %s', (f) =>
    expect(allPass(networkdaysActual, [f]), statuses(networkdaysActual, [f]).join('; ')).toBe(true),
  );

  it('end-of-month terms accepts EOMONTH, DATE and IF versions', () => {
    for (const set of [
      ['=EOMONTH(C2,0)+$I$1', '=MAX(0,$I$2-E2)'],
      ['=DATE(YEAR(C2),MONTH(C2)+1,0)+$I$1', '=IF($I$2>E2,$I$2-E2,0)'],
      ['=EOMONTH(C2,0)+I$1', '=MAX($I$2-E2,0)'],
      ['=EOMONTH(C2,0)+$I$1', '=MAX(0,DAYS($I$2,E2))'],
    ]) {
      expect(allPass(eomonthTerms, set), statuses(eomonthTerms, set).join('; ')).toBe(true);
    }
  });

  it('end-of-month terms flags typed terms and a typed report date', () => {
    expect(allPass(eomonthTerms, ['=EOMONTH(C2,0)+30'])).toBe(false);
    expect(allPass(eomonthTerms, ['=MAX(0,DATE(2026,10,5)-E2)'])).toBe(false);
  });
});
