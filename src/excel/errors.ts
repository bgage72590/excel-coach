/** An error with a message written for the learner, not a stack trace. */
export class CoachError extends Error {
  constructor(
    readonly userMessage: string,
    readonly code = 'coach',
  ) {
    super(userMessage);
  }
}

interface OfficeLikeError {
  code?: string;
  message?: string;
  debugInfo?: { code?: string; message?: string };
}

/** Turns anything thrown by Office.js into a short, actionable sentence. */
export function friendlyError(err: unknown): string {
  if (err instanceof CoachError) return err.userMessage;
  const e = (err ?? {}) as OfficeLikeError;
  const code = e.code ?? e.debugInfo?.code ?? '';
  switch (code) {
    case 'InvalidOperationInCellEditMode':
      return 'Excel is still editing a cell. Press Return or Esc in the sheet, then try again.';
    case 'ItemAlreadyExists':
      return 'Something with that name already exists in this workbook. Try a fresh workbook for practice.';
    case 'ItemNotFound':
      return 'Excel couldn’t find part of the practice sheet. Set it up again with New data.';
    case 'AccessDenied':
    case 'InsertDeleteConflict':
      return 'Excel blocked the change, possibly because the sheet or workbook is protected.';
    case 'ApiNotFound':
      return 'This version of Excel doesn’t support a feature the coach needs. Update Excel and try again.';
    default:
      return e.message ? `Excel couldn’t finish that. ${e.message}` : 'Excel couldn’t finish that. Try again.';
  }
}
