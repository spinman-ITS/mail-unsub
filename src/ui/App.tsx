import { AlertCircle, CheckCircle2, Loader2, Mail, ShieldCheck, Undo2 } from "lucide-react";
import { useEffect, useRef, useState } from "react";
import { findBodyUnsubscribeLinks } from "../lib/bodyLinks";
import { getCurrentMessageRestId, moveMessageToDeletedItems } from "../lib/graph";
import { parseUnsubscribeHeaders, type UnsubscribeHeaders } from "../lib/headerParser";
import {
  fetchAppConfig,
  performBodyLinkUnsubscribe,
  performOneClickUnsubscribe,
  recordUnsubscribedSender
} from "../lib/unsubscribeClient";
import { waitForOutlookMessageContext } from "../office/host";
import { readCurrentMessageInfo } from "../office/messageInfo";
import { readCurrentMessageBodyHtml, readCurrentMessageHeaders } from "../office/readHeaders";
import { ScanPanel } from "./ScanPanel";

type UiState = "booting" | "preview" | "reading" | "submitting" | "success" | "warning" | "error";

type Notice = {
  title: string;
  body: string;
};

type View = "message" | "scan";

export function App() {
  const [isOutlookMessage, setIsOutlookMessage] = useState<boolean | null>(null);
  const [state, setState] = useState<UiState>("booting");
  const [view, setView] = useState<View>("message");
  const [aadClientId, setAadClientId] = useState<string | null>(null);
  // The auto-run unsubscribe closure starts before state updates land, so it reads the ref.
  const aadClientIdRef = useRef<string | null>(null);
  const [headers, setHeaders] = useState<UnsubscribeHeaders | null>(null);
  const [notice, setNotice] = useState<Notice>({
    title: "Preparing unsubscribe",
    body: "Connecting to Outlook and reading the selected message."
  });

  useEffect(() => {
    let cancelled = false;

    async function run() {
      const [hasContext, config] = await Promise.all([waitForOutlookMessageContext(), fetchAppConfig()]);
      if (cancelled) {
        return;
      }

      aadClientIdRef.current = config.aadClientId;
      setAadClientId(config.aadClientId);
      setIsOutlookMessage(hasContext);

      if (!hasContext) {
        setState("preview");
        setNotice({
          title: "Local preview mode",
          body: "The add-in web page is running. Open it from the Outlook Unsubscribe command to process the selected email."
        });
        return;
      }

      await inspectAndUnsubscribe();
    }

    run();

    return () => {
      cancelled = true;
    };
  }, []);

  async function inspectAndUnsubscribe() {
    setState("reading");
    setNotice({ title: "Checking this email", body: "Looking for a standards-based unsubscribe method." });

    try {
      const rawHeaders = await readCurrentMessageHeaders();
      const parsed = parseUnsubscribeHeaders(rawHeaders);
      setHeaders(parsed);

      if (parsed.oneClick && parsed.httpsUrls.length > 0) {
        await submitUnsubscribe(parsed.httpsUrls[0], "one-click");
        return;
      }

      if (parsed.mailtoUrls.length > 0) {
        openMailto(parsed.mailtoUrls[0]);
        return;
      }

      const bodyHtml = await readCurrentMessageBodyHtml();
      const bodyLinks = findBodyUnsubscribeLinks(bodyHtml);
      const bodyHttpsUrl = bodyLinks.find((link) => link.url.toLowerCase().startsWith("https://"))?.url;
      const bodyMailtoUrl = bodyLinks.find((link) => link.url.toLowerCase().startsWith("mailto:"))?.url;

      if (bodyHttpsUrl) {
        await submitUnsubscribe(bodyHttpsUrl, "body-link");
        return;
      }

      if (bodyMailtoUrl) {
        openMailto(bodyMailtoUrl);
        return;
      }

      setState("warning");
      setNotice({
        title: "No supported unsubscribe method found",
        body: "This message does not include a usable List-Unsubscribe header or unsubscribe link."
      });
    } catch (error) {
      setState("error");
      setNotice({
        title: "Could not inspect this message",
        body: error instanceof Error ? error.message : "Outlook did not return message headers."
      });
    }
  }

  async function submitUnsubscribe(url: string, method: "one-click" | "body-link") {
    setState("submitting");
    setNotice({
      title: "Unsubscribing",
      body:
        method === "one-click"
          ? "Sending the one-click unsubscribe request for this email."
          : "Opening the unsubscribe link from the email footer."
    });

    try {
      const response =
        method === "one-click"
          ? await performOneClickUnsubscribe(url)
          : await performBodyLinkUnsubscribe(url, readCurrentMessageInfo()?.userEmail);

      if (!response.ok) {
        setState("error");
        setNotice({ title: "Unsubscribe request failed", body: response.message });
        return;
      }

      const cleanup = await recordAndMoveToDeleted(method);
      setState("success");
      setNotice({
        title: "Unsubscribe request sent",
        body: `${response.message}${cleanup}`
      });
    } catch {
      setState("error");
      setNotice({
        title: "Unsubscribe service is unavailable",
        body: "The unsubscribe API did not respond. Try again in a moment."
      });
    }
  }

  async function recordAndMoveToDeleted(method: string): Promise<string> {
    const info = readCurrentMessageInfo();
    if (info) {
      await recordUnsubscribedSender(info.userEmail, info.senderAddress, method);
    }

    const clientId = aadClientIdRef.current;
    if (!clientId) {
      return " Set up Microsoft Graph access to move handled emails to Deleted Items.";
    }

    try {
      const restId = getCurrentMessageRestId();
      await moveMessageToDeletedItems(clientId, restId);
      return " The email was moved to Deleted Items.";
    } catch {
      return " The email could not be moved to Deleted Items automatically.";
    }
  }

  function openMailto(url: string) {
    const info = readCurrentMessageInfo();
    if (info) {
      recordUnsubscribedSender(info.userEmail, info.senderAddress, "mailto");
    }

    setState("warning");
    setNotice({
      title: "Opening unsubscribe email",
      body: "This sender uses an email-based unsubscribe fallback, so Outlook may open a prefilled message."
    });
    window.setTimeout(() => {
      window.location.href = url;
    }, 500);
  }

  const isBusy = state === "reading" || state === "submitting";
  const showRetry = state === "warning" || state === "error";
  const userEmail = readCurrentMessageInfo()?.userEmail ?? "";

  return (
    <main className="shell">
      <section className="masthead">
        <div className="mark" aria-hidden="true">
          <ShieldCheck size={28} />
        </div>
        <div>
          <h1>Unsubscribe</h1>
          <p>Check the selected message for a safe unsubscribe method.</p>
        </div>
      </section>

      {isOutlookMessage ? (
        <nav className="tabs" aria-label="Add-in views">
          <button className={view === "message" ? "tab active" : "tab"} onClick={() => setView("message")}>
            This email
          </button>
          <button className={view === "scan" ? "tab active" : "tab"} onClick={() => setView("scan")}>
            Scan folder
          </button>
        </nav>
      ) : null}

      {view === "message" ? (
        <>
          <StatusPanel state={state} notice={notice} />

          {showRetry ? (
            <section className="actions" aria-label="Unsubscribe actions">
              <button className="quiet" onClick={inspectAndUnsubscribe} disabled={isBusy || !isOutlookMessage}>
                <Undo2 size={18} />
                Try again
              </button>
            </section>
          ) : null}

          <section className="details" aria-label="Detected unsubscribe methods">
            <h2>Detected methods</h2>
            <dl>
              <div>
                <dt>One-click</dt>
                <dd>{headers ? (headers.oneClick ? "Available" : "Not found") : state === "preview" ? "Open in Outlook" : "Checking"}</dd>
              </div>
              <div>
                <dt>HTTPS</dt>
                <dd>{headers ? (headers.httpsUrls.length ? `${headers.httpsUrls.length} found` : "None") : state === "preview" ? "Open in Outlook" : "Checking"}</dd>
              </div>
              <div>
                <dt>Mailto</dt>
                <dd>{headers ? (headers.mailtoUrls.length ? `${headers.mailtoUrls.length} found` : "None") : state === "preview" ? "Open in Outlook" : "Checking"}</dd>
              </div>
            </dl>
          </section>
        </>
      ) : (
        <ScanPanel aadClientId={aadClientId} userEmail={userEmail} />
      )}
    </main>
  );
}

function StatusPanel({ state, notice }: { state: UiState; notice: Notice }) {
  const icon =
    state === "reading" || state === "submitting" || state === "booting" ? (
      <Loader2 className="spin" size={22} />
    ) : state === "success" ? (
      <CheckCircle2 size={22} />
    ) : state === "warning" ? (
      <Mail size={22} />
    ) : state === "error" ? (
      <AlertCircle size={22} />
    ) : (
      <ShieldCheck size={22} />
    );

  return (
    <section className={`status status-${state}`} aria-live="polite">
      <div className="status-icon">{icon}</div>
      <div>
        <h2>{notice.title}</h2>
        <p>{notice.body}</p>
      </div>
    </section>
  );
}
