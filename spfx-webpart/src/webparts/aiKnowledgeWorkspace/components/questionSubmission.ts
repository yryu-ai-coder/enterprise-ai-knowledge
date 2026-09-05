export interface IQuestionEnterEvent {
  key: string;
  shiftKey: boolean;
  isComposing?: boolean;
  keyCode?: number;
  question: string;
  isLoading: boolean;
}

/**
 * Korean/Japanese/Chinese IMEs emit an Enter key to finalize the active syllable.
 * That key must not send a partial controlled-textarea value to Ask AI.
 */
export function shouldSubmitQuestionOnEnter(event: IQuestionEnterEvent): boolean {
  return event.key === 'Enter'
    && !event.shiftKey
    && !event.isComposing
    && event.keyCode !== 229
    && Boolean(event.question.trim())
    && !event.isLoading;
}

export function shouldDeferQuestionSubmitForComposition(event: IQuestionEnterEvent): boolean {
  return event.key === 'Enter'
    && !event.shiftKey
    && (Boolean(event.isComposing) || event.keyCode === 229)
    && Boolean(event.question.trim())
    && !event.isLoading;
}
