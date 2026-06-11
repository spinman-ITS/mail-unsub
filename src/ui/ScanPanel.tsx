import { CheckCircle2, Loader2, Mail, Search, Trash2 } from "lucide-react";
import { useState } from "react";
import { getCurrentMessageRestId, getParentFolderId, listFolderMessages, moveMessageToDeletedItems } from "../lib/graph";
import { findScanCandidates, type ScanCandidate } from "../lib/scan";
import {
  fetchUnsubscribedSenders,
  performBodyLinkUnsubscribe,
  performOneClickUnsubscribe,
  recordUnsubscribedSender
} from "../lib/unsubscribeClient";

type RowStatus = "idle" | "working" | "done" | "failed" | "mailto";

type RowState = { status: RowStatus; note?: string };

type ScanState = "idle" | "scanning" | "ready" | "error";

export function ScanPanel({ aadClientId, userEmail }: { aadClientId: string | null; userEmail: string }) {
  const [state, setState] = useState<ScanState>("idle");
  const [message, setMessage] = useState<string>("");
  const [candidates, setCandidates] = useState<ScanCandidate[]>([]);
  const [rowStates, setRowStates] = useState<Record<string, RowState>>({});

  function setRow(messageId: string, status: RowStatus, note?: string) {
    setRowStates((current) => ({ ...current, [messageId]: { status, note } }));
  }

  if (!aadClientId) {
    return (
      <section className="scan" aria-label="Folder scan">
        <p className="scan-empty">
          Folder scan needs Microsoft Graph access. Set the AAD_CLIENT_ID environment variable on the server after
          registering the add-in in Microsoft Entra (see README).
        </p>
      </section>
    );
  }

  async function runScan() {
    setState("scanning");
    setMessage("Reading this folder and your unsubscribe history…");

    try {
      const restId = getCurrentMessageRestId();
      const folderId = await getParentFolderId(aadClientId!, restId);
      const [messages, unsubscribed] = await Promise.all([
        listFolderMessages(aadClientId!, folderId),
        fetchUnsubscribedSenders(userEmail)
      ]);

      const found = findScanCandidates(
        messages,
        unsubscribed.map((sender) => sender.senderAddress)
      );

      setCandidates(found);
      setRowStates({});
      setState("ready");
      setMessage(
        found.length === 0
          ? `Scanned ${messages.length} recent messages. No new senders with unsubscribe support were found.`
          : `Found ${found.length} sender${found.length === 1 ? "" : "s"} you can unsubscribe from (${messages.length} messages scanned).`
      );
    } catch (error) {
      setState("error");
      setMessage(error instanceof Error ? error.message : "The folder scan failed.");
    }
  }

  async function unsubscribeCandidate(candidate: ScanCandidate) {
    setRow(candidate.messageId, "working");

    const httpsUrl = candidate.headers.httpsUrls[0];
    const mailtoUrl = candidate.headers.mailtoUrls[0];

    try {
      if (httpsUrl) {
        const response = candidate.headers.oneClick
          ? await performOneClickUnsubscribe(httpsUrl)
          : await performBodyLinkUnsubscribe(httpsUrl, userEmail);

        if (!response.ok) {
          setRow(candidate.messageId, "failed", response.message);
          return;
        }

        await recordUnsubscribedSender(userEmail, candidate.senderAddress, candidate.headers.oneClick ? "one-click" : "https-link");
        const moved = await moveMessageToDeletedItems(aadClientId!, candidate.messageId)
          .then(() => true)
          .catch(() => false);
        setRow(
          candidate.messageId,
          "done",
          moved ? "Unsubscribed and moved to Deleted Items." : "Unsubscribed, but the email could not be deleted."
        );
        return;
      }

      if (mailtoUrl) {
        await recordUnsubscribedSender(userEmail, candidate.senderAddress, "mailto");
        setRow(candidate.messageId, "mailto", "This sender uses email-based unsubscribe; an email was opened.");
        window.location.href = mailtoUrl;
        return;
      }

      setRow(candidate.messageId, "failed", "No usable unsubscribe method.");
    } catch {
      setRow(candidate.messageId, "failed", "The unsubscribe request failed.");
    }
  }

  return (
    <section className="scan" aria-label="Folder scan">
      <button className="primary" onClick={runScan} disabled={state === "scanning"}>
        {state === "scanning" ? <Loader2 className="spin" size={18} /> : <Search size={18} />}
        Scan this folder
      </button>

      {message ? <p className={state === "error" ? "scan-error" : "scan-note"}>{message}</p> : null}

      {candidates.length > 0 ? (
        <ul className="scan-list">
          {candidates.map((candidate) => {
            const row = rowStates[candidate.messageId] ?? { status: "idle" as RowStatus };
            return (
              <li key={candidate.messageId}>
                <div className="scan-row">
                  <div className="scan-sender">
                    <strong>{candidate.senderName}</strong>
                    <span>{candidate.senderAddress}</span>
                    <span className="scan-subject">{candidate.subject}</span>
                  </div>
                  <div className="scan-action">
                    {row.status === "idle" ? (
                      <button className="quiet" onClick={() => unsubscribeCandidate(candidate)}>
                        <Trash2 size={16} />
                        Unsubscribe
                      </button>
                    ) : row.status === "working" ? (
                      <Loader2 className="spin" size={18} />
                    ) : row.status === "done" ? (
                      <span className="scan-done">
                        <CheckCircle2 size={16} /> Unsubscribed
                      </span>
                    ) : row.status === "mailto" ? (
                      <span className="scan-done">
                        <Mail size={16} /> Email opened
                      </span>
                    ) : (
                      <span className="scan-failed">Failed</span>
                    )}
                  </div>
                </div>
                {row.note ? (
                  <p className={row.status === "failed" ? "scan-row-note scan-error" : "scan-row-note"}>{row.note}</p>
                ) : null}
              </li>
            );
          })}
        </ul>
      ) : null}
    </section>
  );
}
