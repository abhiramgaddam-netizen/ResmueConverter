/**
 * Extracts the first name only and formats it with the first letter capitalized and subsequent letters lowercase.
 * Skips single-letter initials (e.g., "K. Shekar" -> "Shekar") and honorific titles (e.g., "Dr. John" -> "John").
 */
export function formatCandidateFirstName(fullName: string): string {
  if (!fullName || !fullName.trim()) return 'Candidate';
  const cleaned = fullName.trim().replace(/^['"]|['"]$/g, '');
  const tokens = cleaned.split(/\s+/).filter(Boolean);

  let chosen = tokens[0] || 'Candidate';
  for (const token of tokens) {
    const stripped = token.replace(/[^a-zA-Z]/g, '');
    if (stripped.length <= 1) continue;
    if (/^(mr|mrs|ms|dr|prof)\.?$/i.test(token)) continue;
    chosen = token;
    break;
  }

  return chosen
    .split('-')
    .map((part) => {
      const lettersOnly = part.replace(/[^a-zA-Z]/g, '');
      if (!lettersOnly) return '';
      return lettersOnly.charAt(0).toUpperCase() + lettersOnly.slice(1).toLowerCase();
    })
    .filter(Boolean)
    .join('-');
}
