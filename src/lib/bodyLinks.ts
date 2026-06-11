export type BodyUnsubscribeLink = {
  url: string;
  label: string;
};

const unsubscribeWords =
  /\b(unsubscribe|opt[\s-]?out|email preferences|subscription preferences|manage preferences|preference center|communication preferences|update (?:your |my )?preferences)\b/i;

export function findBodyUnsubscribeLinks(htmlOrText: string): BodyUnsubscribeLink[] {
  const doc = new DOMParser().parseFromString(htmlOrText, "text/html");
  const candidates: BodyUnsubscribeLink[] = [];

  for (const anchor of Array.from(doc.querySelectorAll("a[href]"))) {
    const href = anchor.getAttribute("href")?.trim();
    if (!href || !isSupportedHref(href)) {
      continue;
    }

    const label = normalizeWhitespace(anchor.textContent || href);
    const ownText = `${label} ${href}`;
    const genericText = /^(click here|here|manage|update)$/i.test(label);
    const nearbyText = genericText ? normalizeWhitespace(anchor.parentElement?.textContent || "") : "";
    const combined = `${ownText} ${nearbyText}`;

    const hrefLower = href.toLowerCase();
    if (unsubscribeWords.test(combined) || hrefLower.includes("unsubscribe") || hrefLower.includes("preference")) {
      candidates.push({ url: href, label: label || href });
    }
  }

  // Explicit unsubscribe links beat preference-page links: callers use the first match.
  const score = (link: BodyUnsubscribeLink) => (`${link.url} ${link.label}`.toLowerCase().includes("unsubscribe") ? 0 : 1);
  return uniqueByUrl(candidates).sort((a, b) => score(a) - score(b));
}

function isSupportedHref(href: string): boolean {
  const lower = href.toLowerCase();
  return lower.startsWith("https://") || lower.startsWith("mailto:");
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function uniqueByUrl(links: BodyUnsubscribeLink[]): BodyUnsubscribeLink[] {
  const seen = new Set<string>();
  return links.filter((link) => {
    if (seen.has(link.url)) {
      return false;
    }
    seen.add(link.url);
    return true;
  });
}
