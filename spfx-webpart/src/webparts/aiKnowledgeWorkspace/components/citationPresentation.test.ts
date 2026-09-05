import { formatCitationSnippet } from './citationPresentation';

describe('citation presentation', () => {
  it('collapses extracted PDF whitespace and bounds long source excerpts', () => {
    const raw = `Name: Example\n=============================\n${'Long extracted source text '.repeat(30)}`;
    const formatted = formatCitationSnippet(raw, 120);

    expect(formatted).toBe('Name: Example ============================= Long extracted source text Long extracted source text Long extracted source…');
    expect(formatted.length).toBeLessThanOrEqual(121);
  });
});
