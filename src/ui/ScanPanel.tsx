import { CheckCircle2, Loader2, Mail, Search, Trash2 } from "lucide-react";
import { useMemo, useState } from "react";
import { getCurrentMessageRestId, getParentFolderId, listFolderMessages, moveMessageToDeletedItems } from "../lib/graph";
import { findScanCandidates, type ScanCandidate } from "../lib/scan";
import {
  fetchUnsubscribedSenders,
  performBodyLinkUnsubscribe,
  performOneClickUnsubscribe,
  recordUnsubscribedSender
} from "../lib/unsubscribeClient";
import { openDialogAndWait } from "../office/dialog";

type RowStatus = "idle" | "working" | "done" | "failed" | "skipped" | "mailto";

type RowState = { status: RowStatus; note?: string };

type ScanState = "idle" | "scanning" | "ready" | "error";

export function ScanPanel({ aadClientId, userEmail }: { aadClientId: string | null; userEmail: string }) {
  const [state, setState] = useState<ScanState>("idle");
  const [message, setMessage] = useState<string>("");
  const [candidates, setCandidates] = useState<ScanCandidate[]>([]);
  const [rowStates, setRowStates] = useState<Record<string, RowState>>({});
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [bulkRunning, setBulkRunning] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<{ done: number; total: number } | null>(null);

  function setRow(messageId: string, status: RowStatus, note?: string) {
    setRowStates((current) => ({ ...current, [messageId]: { status, note } }));
  }

  const selectableIds = useMemo(() => {
    return candidates
      .filter((c) => (rowStates[c.messageId]?.status ?? "idle") === "idle")
      .map((c) => c.messageId);
  }, [candidates, rowStates]);

  const allSelected = selectableIds.length > 0 && selectableIds.every((id) => selected.has(id));
  const selectedCount = selectableIds.filter((id) => selected.has(id)).length;

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
    setSelected(new Set());
    setBulkProgress(null);

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
          : `Found ${found.length} sender${found.length === 1 ? "" : "s"} you can unsubscribe from (${messages.length} messages scanned). Select the ones you want to handle.`
      );
    } catch (error) {
      setState("error");
      setMessage(error instanceof Error ? error.message : "The folder scan failed.");
    }
  }

  function toggleSelection(messageId: string) {
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(messageId)) {
        next.delete(messageId);
      } else {
        next.add(messageId);
      }
      return next;
    });
  }

  function toggleSelectAll() {
    setSelected((current) => {
      if (selectableIds.every((id) => current.has(id))) {
        const next = new Set(current);
        for (const id of selectableIds) {
          next.delete(id);
        }
        return next;
      }
      return new Set(selectableIds);
    });
  }

  async function unsubscribeOne(candidate: ScanCandidate) {
    setRow(candidate.messageId, "working");

    const httpsUrl = candidate.headers.httpsUrls[0];
    const mailtoUrl = candidate.headers.mailtoUrls[0];

    try {
      if (httpsUrl) {
        const response = candidate.headers.oneClick
          ? await performOneClickUnsubscribe(httpsUrl)
          : await performBodyLinkUnsubscribe(httpsUrl);

        if (!response.ok) {
          setRow(candidate.messageId, "failed", response.message);
          return;
        }

        if (response.requiresBrowser && response.finalUrl) {
          setRow(candidate.messageId, "working", "Opening preference page — complete it in the dialog.");
          await openDialogAndWait(response.finalUrl);
        }

        await recordUnsubscribedSender(
          userEmail,
          candidate.senderAddress,
          candidate.headers.oneClick ? "one-click" : "https-link"
        );
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
        setRow(
          candidate.messageId,
          "mailto",
          "This sender uses email-based unsubscribe. Use the per-row Open button to send the unsubscribe email."
        );
        return;
      }

      setRow(candidate.messageId, "failed", "No usable unsubscribe method.");
    } catch {
      setRow(candidate.messageId, "failed", "The unsubscribe request failed.");
    }
  }

  function openMailtoFor(candidate: ScanCandidate) {
    const mailtoUrl = candidate.headers.mailtoUrls[0];
    if (!mailtoUrl) return;
    recordUnsubscribedSender(userEmail, candidate.senderAddress, "mailto");
    setRow(candidate.messageId, "mailto", "Email-based unsubscribe was opened.");
    window.location.href = mailtoUrl;
  }

  async function bulkUnsubscribe() {
    const ids = selectableIds.filter((id) => selected.has(id));
    if (ids.length === 0) return;

    setBulkRunning(true);
    setBulkProgress({ done: 0, total: ids.length });

    for (let i = 0; i < ids.length; i++) {
      const candidate = candidates.find((c) => c.messageId === ids[i]);
      if (candidate) {
        await unsubscribeOne(candidate);
      }
      setBulkProgress({ done: i + 1, total: ids.length });
    }

    setSelected(new Set());
    setBulkRunning(false);
  }

  return (
    <section className="scan" aria-label="Folder scan">
      <button className="primary" onClick={runScan} disabled={state === "scanning" || bulkRunning}>
        {state === "scanning" ? <Loader2 className="spin" size={18} /> : <Search size={18} />}
        {state === "ready" ? "Rescan this folder" : "Scan this folder"}
      </button>

      {message ? <p className={state === "error" ? "scan-error" : "scan-note"}>{message}</p> : null}

      {candidates.length > 0 ? (
        <>
          <div className="scan-toolbar">
            <label className="scan-select-all">
              <input
                type="checkbox"
                checked={allSelected}
                onChange={toggleSelectAll}
                disabled={bulkRunning || selectableIds.length === 0}
              />
              {allSelected ? "Clear selection" : "Select all"}
            </label>
            <button
              className="primary"
              onClick={bulkUnsubscribe}
              disabled={bulkRunning || selectedCount === 0}
            >
              {bulkRunning ? <Loader2 className="spin" size={16} /> : <Trash2 size={16} />}
              {bulkRunning && bulkProgress
                ? `Working ${bulkProgress.done}/${bulkProgress.total}`
                : `Unsubscribe ${selectedCount || ""} selected`.trim()}
            </button>
          </div>

          <ul className="scan-list">
            {candidates.map((candidate) => {
              const row = rowStates[candidate.messageId] ?? { status: "idle" as RowStatus };
              const isIdle = row.status === "idle";
              const hasMailto = candidate.headers.mailtoUrls.length > 0;
              const hasHttps = candidate.headers.httpsUrls.length > 0;
              return (
                <li key={candidate.messageId}>
                  <div className="scan-row">
                    <input
                      type="checkbox"
                      className="scan-row-check"
                      checked={selected.has(candidate.messageId)}
                      onChange={() => toggleSelection(candidate.messageId)}
                      disabled={!isIdle || bulkRunning}
                      aria-label={`Select ${candidate.senderName}`}
                    />
                    <div className="scan-sender">
                      <strong>{candidate.senderName}</strong>
                      <span>{candidate.senderAddress}</span>
                      <span className="scan-subject">{candidate.subject}</span>
                    </div>
                    <div className="scan-action">
                      {isIdle ? (
                        hasHttps ? (
                          <button className="quiet" onClick={() => unsubscribeOne(candidate)} disabled={bulkRunning}>
                            <Trash2 size={16} />
                            Unsubscribe
                          </button>
                        ) : hasMailto ? (
                          <button className="quiet" onClick={() => openMailtoFor(candidate)} disabled={bulkRunning}>
                            <Mail size={16} />
                            Open email
                          </button>
                        ) : null
                      ) : row.status === "working" ? (
                        <Loader2 className="spin" size={18} />
                      ) : row.status === "done" ? (
                        <span className="scan-done">
                          <CheckCircle2 size={16} /> Unsubscribed
                        </span>
                      ) : row.status === "mailto" ? (
                        <button className="quiet" onClick={() => openMailtoFor(candidate)} disabled={bulkRunning}>
                          <Mail size={16} />
                          Open email
                        </button>
                      ) : row.status === "skipped" ? (
                        <span className="scan-skipped">Skipped</span>
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
        </>
      ) : null}
    </section>
  );
}
