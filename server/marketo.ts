export type PreferenceCenterTarget = {
  pageUrl: string;
};

// Marketo preference centers embed a formDescriptor JSON literal with a
// checkbox named "Unsubscribed" — the "unsubscribe from all marketing emails"
// toggle. We don't try to submit it automatically (the service hashes the
// payload with a moving target). Detection tells the client to open the page.
export function isPreferenceCenter(html: string): boolean {
  if (!html.includes('"Name":"Unsubscribed"')) {
    return false;
  }
  return extractFormDescriptor(html) !== null;
}

type FormDescriptor = {
  Id?: number;
  Vid?: number;
};

function extractFormDescriptor(html: string): FormDescriptor | null {
  const match = html.match(/var\s+formDescriptor\s*=\s*(\{)/);
  if (!match || match.index === undefined) {
    return null;
  }

  const start = match.index + match[0].length - 1;
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let i = start; i < html.length && i < start + 500_000; i++) {
    const char = html[i];

    if (escaped) {
      escaped = false;
      continue;
    }
    if (inString) {
      if (char === "\\") {
        escaped = true;
      } else if (char === '"') {
        inString = false;
      }
      continue;
    }
    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === "{") {
      depth++;
    } else if (char === "}") {
      depth--;
      if (depth === 0) {
        try {
          return JSON.parse(html.slice(start, i + 1)) as FormDescriptor;
        } catch {
          return null;
        }
      }
    }
  }

  return null;
}
