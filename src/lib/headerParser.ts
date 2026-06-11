export type UnsubscribeHeaders = {
  oneClick: boolean;
  httpsUrls: string[];
  mailtoUrls: string[];
};

type HeaderMap = Map<string, string[]>;

export function parseUnsubscribeHeaders(rawHeaders: string): UnsubscribeHeaders {
  const headers = parseHeaders(rawHeaders);
  const listUnsubscribe = headers.get("list-unsubscribe") ?? [];
  const postValues = headers.get("list-unsubscribe-post") ?? [];
  const uris = listUnsubscribe.flatMap(extractAngleBracketUris);

  return {
    oneClick: postValues.some((value) => /(?:^|;|\s)List-Unsubscribe=One-Click(?:$|;|\s)/i.test(value)),
    httpsUrls: unique(uris.filter((uri) => uri.toLowerCase().startsWith("https://"))),
    mailtoUrls: unique(uris.filter((uri) => uri.toLowerCase().startsWith("mailto:")))
  };
}

function parseHeaders(rawHeaders: string): HeaderMap {
  const unfolded = rawHeaders.replace(/\r?\n[ \t]+/g, " ");
  const map: HeaderMap = new Map();

  for (const line of unfolded.split(/\r?\n/)) {
    const separator = line.indexOf(":");
    if (separator <= 0) {
      continue;
    }

    const name = line.slice(0, separator).trim().toLowerCase();
    const value = line.slice(separator + 1).trim();
    const values = map.get(name) ?? [];
    values.push(value);
    map.set(name, values);
  }

  return map;
}

function extractAngleBracketUris(value: string): string[] {
  return [...value.matchAll(/<([^>]+)>/g)].map((match) => match[1].trim());
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}
