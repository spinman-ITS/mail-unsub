import { isIP } from "node:net";

export type UrlValidationResult =
  | {
      ok: true;
      url: string;
    }
  | {
      ok: false;
      reason: string;
    };

const blockedHosts = new Set(["localhost", "localhost.localdomain"]);

export async function validateUnsubscribeUrl(value: string): Promise<UrlValidationResult> {
  let url: URL;

  try {
    url = new URL(value);
  } catch {
    return { ok: false, reason: "The unsubscribe URL is not valid." };
  }

  if (url.protocol !== "https:") {
    return { ok: false, reason: "Only HTTPS unsubscribe URLs are supported." };
  }

  const hostname = url.hostname.toLowerCase();
  if (blockedHosts.has(hostname) || isPrivateIpLiteral(hostname)) {
    return { ok: false, reason: "Private and local network unsubscribe URLs are not allowed." };
  }

  return { ok: true, url: url.toString() };
}

export function isSafeRedirectTarget(value: string): boolean {
  try {
    const url = new URL(value);
    const isWebProtocol = url.protocol === "https:" || url.protocol === "http:";
    return isWebProtocol && !blockedHosts.has(url.hostname.toLowerCase()) && !isPrivateIpLiteral(url.hostname);
  } catch {
    return false;
  }
}

function isPrivateIpLiteral(hostname: string): boolean {
  const ipVersion = isIP(hostname);
  if (ipVersion === 0) {
    return false;
  }

  if (ipVersion === 6) {
    return hostname === "::1" || hostname.toLowerCase().startsWith("fc") || hostname.toLowerCase().startsWith("fd");
  }

  const parts = hostname.split(".").map(Number);
  const [first, second] = parts;

  return (
    first === 10 ||
    first === 127 ||
    (first === 172 && second >= 16 && second <= 31) ||
    (first === 192 && second === 168) ||
    (first === 169 && second === 254) ||
    first === 0
  );
}
