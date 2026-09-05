/** Converts common model Markdown and encoded text into clean chat UI text. */
function decodeHtmlEntities(text: string): string {
  const entities: Record<string, string> = {
    amp: '&', quot: '"', apos: "'", lt: '<', gt: '>'
  };
  // Decode twice so values that were HTML-escaped before JSON/UI handling,
  // such as &amp;quot;, become their intended character.
  let decoded = text;
  for (let pass = 0; pass < 2; pass += 1) {
    decoded = decoded.replace(/&(amp|quot|apos|lt|gt);/gi, (match, name: string) => entities[name.toLowerCase()] || match);
  }
  return decoded;
}

export function formatAnswerForDisplay(answer: string): string {
  return decodeHtmlEntities(answer)
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/__(.*?)__/g, '$1')
    .replace(/^#{1,6}\s*/gm, '')
    .replace(/^[ \t]*[-*][ \t]+/gm, '• ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
