import type { GraphMessageSummary } from "./graph";
import { parseUnsubscribeHeaders, type UnsubscribeHeaders } from "./headerParser";

export type ScanCandidate = {
  messageId: string;
  subject: string;
  senderName: string;
  senderAddress: string;
  receivedDateTime?: string;
  headers: UnsubscribeHeaders;
};

export function findScanCandidates(
  messages: GraphMessageSummary[],
  alreadyUnsubscribed: Iterable<string>
): ScanCandidate[] {
  const unsubscribed = new Set([...alreadyUnsubscribed].map((address) => address.trim().toLowerCase()));
  const seenSenders = new Set<string>();
  const candidates: ScanCandidate[] = [];

  for (const message of messages) {
    const senderAddress = message.from?.emailAddress?.address?.trim().toLowerCase() ?? "";
    if (!senderAddress || unsubscribed.has(senderAddress) || seenSenders.has(senderAddress)) {
      continue;
    }

    const headers = parseUnsubscribeHeaders(headersToRaw(message.internetMessageHeaders));
    if (headers.httpsUrls.length === 0 && headers.mailtoUrls.length === 0) {
      continue;
    }

    seenSenders.add(senderAddress);
    candidates.push({
      messageId: message.id,
      subject: message.subject?.trim() || "(no subject)",
      senderName: message.from?.emailAddress?.name?.trim() || senderAddress,
      senderAddress,
      receivedDateTime: message.receivedDateTime,
      headers
    });
  }

  return candidates;
}

function headersToRaw(headers: { name: string; value: string }[] | undefined): string {
  if (!headers) {
    return "";
  }
  return headers.map((header) => `${header.name}: ${header.value}`).join("\r\n");
}
