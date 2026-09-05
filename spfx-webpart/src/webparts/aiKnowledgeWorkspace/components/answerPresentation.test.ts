import { formatAnswerForDisplay } from './answerPresentation';

describe('answer presentation', () => {
  it('removes markdown emphasis and turns headings into plain readable text', () => {
    expect(formatAnswerForDisplay('**Short answer:**\n\n- **Yes** — it is mentioned.')).toBe('Short answer:\n\n• Yes — it is mentioned.');
  });

  it('decodes encoded quotation marks in submitted questions', () => {
    expect(formatAnswerForDisplay('&amp;quot;Current market analysis&amp;quot; 에 대해 요약 정리 해줘')).toBe('"Current market analysis" 에 대해 요약 정리 해줘');
  });
});
