import express from "express";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { addAuditEntry, getAuditEntries, redactUrl } from "./auditLog";
import { createSenderStore, emailDomain, normalizeEmail } from "./db";
import { isPreferenceCenter } from "./marketo";
import { isSafeRedirectTarget, validateUnsubscribeUrl } from "./safety";

const app = express();
const port = Number(process.env.PORT || 8787);
const isLocalDev = process.env.NODE_ENV === "development";

const senderStore = await createSenderStore();

app.use(express.json({ limit: "16kb" }));

app.use((request, response, next) => {
  const origin = request.header("origin");
  const allowed =
    origin &&
    (origin.startsWith("https://localhost:") ||
      origin.startsWith("http://localhost:") ||
      (process.env.ALLOWED_ORIGINS ?? "").split(",").map((value) => value.trim()).includes(origin));

  if (allowed) {
    response.header("Access-Control-Allow-Origin", origin);
    response.header("Vary", "Origin");
  }
  response.header("Access-Control-Allow-Headers", "Content-Type");
  response.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  if (request.method === "OPTIONS") {
    response.sendStatus(204);
    return;
  }
  next();
});

app.post("/api/unsubscribe", async (request, response) => {
  const url = typeof request.body?.url === "string" ? request.body.url : "";
  await submitUnsubscribeRequest(url, "POST", response);
});

app.post("/api/unsubscribe-link", async (request, response) => {
  const url = typeof request.body?.url === "string" ? request.body.url : "";
  await submitUnsubscribeRequest(url, "GET", response);
});

app.get("/health", (_request, response) => {
  response.json({ ok: true });
});

app.get("/api/config", (_request, response) => {
  response.json({
    aadClientId: process.env.AAD_CLIENT_ID || null
  });
});

app.post("/api/unsubscribed", async (request, response) => {
  const userEmail = typeof request.body?.userEmail === "string" ? normalizeEmail(request.body.userEmail) : "";
  const senderAddress = typeof request.body?.senderAddress === "string" ? normalizeEmail(request.body.senderAddress) : "";
  const method = typeof request.body?.method === "string" ? request.body.method : "unknown";

  if (!userEmail.includes("@") || !senderAddress.includes("@")) {
    response.status(400).json({ ok: false, message: "userEmail and senderAddress are required." });
    return;
  }

  await senderStore.record({
    userEmail,
    senderAddress,
    senderDomain: emailDomain(senderAddress),
    method
  });
  response.json({ ok: true });
});

app.get("/api/unsubscribed", async (request, response) => {
  const userEmail = typeof request.query.userEmail === "string" ? normalizeEmail(request.query.userEmail) : "";

  if (!userEmail.includes("@")) {
    response.status(400).json({ ok: false, message: "userEmail is required." });
    return;
  }

  const senders = await senderStore.listForUser(userEmail);
  response.json({ ok: true, senders });
});

app.get("/api/logs", (_request, response) => {
  response.json({ entries: getAuditEntries() });
});

const distDir = resolve(dirname(fileURLToPath(import.meta.url)), "../dist");
if (!isLocalDev && existsSync(distDir)) {
  app.use(express.static(distDir));
  app.get("/", (_request, response) => {
    response.sendFile(resolve(distDir, "taskpane.html"));
  });
}

if (isLocalDev) {
  const { createServer } = await import("node:https");
  const { getHttpsServerOptions } = await import("office-addin-dev-certs");
  const httpsOptions = await getHttpsServerOptions(365);
  createServer(httpsOptions, app).listen(port, () => {
    console.log(`Unsubscribe API listening on https://localhost:${port}`);
  });
} else {
  app.listen(port, () => {
    console.log(`Unsubscribe service listening on port ${port}`);
  });
}

async function submitUnsubscribeRequest(
  url: string,
  method: "GET" | "POST",
  response: express.Response
) {
  const validation = await validateUnsubscribeUrl(url);

  if (!validation.ok) {
    addAuditEntry({
      method,
      target: redactUrl(url),
      outcome: "blocked",
      message: validation.reason
    });
    response.status(400).json({ ok: false, message: validation.reason });
    return;
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  const maxRedirects = 3;

  try {
    let currentUrl = validation.url;
    let finalResponse: Response | null = null;

    for (let hop = 0; hop <= maxRedirects; hop++) {
      const hopResponse = await fetch(currentUrl, {
        method: hop === 0 ? method : "GET",
        redirect: "manual",
        signal: controller.signal,
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "User-Agent": "Outlook-Unsubscribe-Addin/0.1"
        },
        body: hop === 0 && method === "POST" ? "List-Unsubscribe=One-Click" : undefined
      });

      const location = hopResponse.headers.get("location");
      const isRedirect = hopResponse.status >= 300 && hopResponse.status < 400 && location;

      if (!isRedirect) {
        finalResponse = hopResponse;
        break;
      }

      const resolvedLocation = new URL(location, currentUrl).toString();
      if (!isSafeRedirectTarget(resolvedLocation)) {
        addAuditEntry({
          method,
          target: redactUrl(validation.url),
          outcome: "blocked",
          status: hopResponse.status,
          message: "Unsafe redirect was not followed.",
          redirectTarget: redactUrl(resolvedLocation)
        });
        response.status(502).json({
          ok: false,
          status: hopResponse.status,
          message: "The sender returned an unsafe redirect, so it was not followed."
        });
        return;
      }

      currentUrl = resolvedLocation;
    }

    if (!finalResponse) {
      addAuditEntry({
        method,
        target: redactUrl(validation.url),
        outcome: "failed",
        message: `More than ${maxRedirects} redirects; gave up.`
      });
      response.status(502).json({
        ok: false,
        message: "The unsubscribe link redirected too many times."
      });
      return;
    }

    if (finalResponse.status >= 200 && finalResponse.status < 300) {
      const preferencePageUrl =
        method === "GET" ? await detectPreferenceCenter(finalResponse, currentUrl) : null;

      addAuditEntry({
        method,
        target: redactUrl(validation.url),
        outcome: "accepted",
        status: finalResponse.status,
        message: preferencePageUrl
          ? "Preference center detected, opening for user."
          : method === "POST"
            ? "One-click unsubscribe accepted."
            : "Unsubscribe link opened.",
        redirectTarget: currentUrl === validation.url ? undefined : redactUrl(currentUrl)
      });

      if (preferencePageUrl) {
        response.json({
          ok: true,
          requiresBrowser: true,
          finalUrl: preferencePageUrl,
          status: finalResponse.status,
          message:
            "This sender uses an email preference page. Check \"Unsubscribe from all\" and submit the form in the window that opens."
        });
        return;
      }

      response.json({
        ok: true,
        status: finalResponse.status,
        message:
          method === "POST"
            ? "The sender accepted the one-click unsubscribe request."
            : "The unsubscribe link was opened successfully."
      });
      return;
    }

    addAuditEntry({
      method,
      target: redactUrl(validation.url),
      outcome: "failed",
      status: finalResponse.status,
      message: `Sender returned HTTP ${finalResponse.status}.`,
      redirectTarget: currentUrl === validation.url ? undefined : redactUrl(currentUrl)
    });
    response.status(502).json({
      ok: false,
      status: finalResponse.status,
      message: `The sender returned HTTP ${finalResponse.status}.`
    });
  } catch (error) {
    const timedOut = error instanceof Error && error.name === "AbortError";
    addAuditEntry({
      method,
      target: redactUrl(validation.url),
      outcome: "failed",
      message: timedOut ? "Request timed out." : "Request failed before completion."
    });
    response.status(504).json({
      ok: false,
      message: timedOut ? "The unsubscribe request timed out." : "The unsubscribe request failed."
    });
  } finally {
    clearTimeout(timeout);
  }
}

async function detectPreferenceCenter(pageResponse: Response, pageUrl: string): Promise<string | null> {
  try {
    const contentType = pageResponse.headers.get("content-type") ?? "";
    if (!contentType.includes("html")) {
      return null;
    }
    const html = (await pageResponse.text()).slice(0, 1_000_000);
    return isPreferenceCenter(html) ? pageUrl : null;
  } catch {
    return null;
  }
}
