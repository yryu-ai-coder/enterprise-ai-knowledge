import { formatAnswerForDisplay } from './answerPresentation';

/** Keeps citation cards compact and prevents raw extraction text from widening the drawer. */
export function formatCitationSnippet(snippet: string, maxChars = 320): string {
  const normalized = formatAnswerForDisplay(snippet).replace(/\s+/g, ' ').trim();
  if (normalized.length <= maxChars) return normalized;
  return `${normalized.slice(0, maxChars).replace(/\s+$/, '')}…`;
}
