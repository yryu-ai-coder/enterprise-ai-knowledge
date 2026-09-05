import { shouldDeferQuestionSubmitForComposition, shouldSubmitQuestionOnEnter } from './questionSubmission';

describe('shouldSubmitQuestionOnEnter', () => {
  it('does not submit Enter while a Korean IME composition is still active', () => {
    expect(shouldSubmitQuestionOnEnter({
      key: 'Enter',
      shiftKey: false,
      isComposing: true,
      keyCode: 229,
      question: 'Actual Expenses를 물어 봤는데',
      isLoading: false
    })).toBe(false);
  });

  it('does not submit the IME process key even when a browser omits isComposing', () => {
    expect(shouldSubmitQuestionOnEnter({
      key: 'Enter',
      shiftKey: false,
      isComposing: false,
      keyCode: 229,
      question: 'Actual Expenses를 물어 봤는데',
      isLoading: false
    })).toBe(false);
  });

  it('defers a Korean IME Enter so the completed syllable is included before sending', () => {
    expect(shouldDeferQuestionSubmitForComposition({
      key: 'Enter',
      shiftKey: false,
      isComposing: true,
      keyCode: 229,
      question: 'Actual Expenses를 물어 봤는데',
      isLoading: false
    })).toBe(true);
  });

  it('submits a completed non-empty question on plain Enter', () => {
    expect(shouldSubmitQuestionOnEnter({
      key: 'Enter',
      shiftKey: false,
      isComposing: false,
      keyCode: 13,
      question: 'Actual Expenses를 물어 봤는데',
      isLoading: false
    })).toBe(true);
  });
});
