import express from "express";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { addAuditEntry, getAuditEntries, redactUrl } from "./auditLog";
import { createSenderStore, emailDomain, normalizeEmail } from "./db";
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

async function submitUnsubscribeRequest(url: string, method: "GET" | "POST", response: express.Response) {
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
  const timeout = setTimeout(() => controller.abort(), 8000);

  try {
    const unsubscribeResponse = await fetch(validation.url, {
      method,
      redirect: "manual",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": "Outlook-Unsubscribe-Addin/0.1"
      },
      body: method === "POST" ? "List-Unsubscribe=One-Click" : undefined
    });

    const location = unsubscribeResponse.headers.get("location");
    const resolvedLocation = location ? new URL(location, validation.url).toString() : undefined;
    if (resolvedLocation && !isSafeRedirectTarget(resolvedLocation)) {
      addAuditEntry({
        method,
        target: redactUrl(validation.url),
        outcome: "blocked",
        status: unsubscribeResponse.status,
        message: "Unsafe redirect was not followed.",
        redirectTarget: redactUrl(resolvedLocation)
      });
      response.status(502).json({
        ok: false,
        status: unsubscribeResponse.status,
        message: "The sender returned an unsafe redirect, so it was not followed."
      });
      return;
    }

    if (resolvedLocation) {
      const redirectedResponse = await fetch(resolvedLocation, {
        method: "GET",
        redirect: "manual",
        signal: controller.signal,
        headers: {
          "User-Agent": "Outlook-Unsubscribe-Addin/0.1"
        }
      });

      addAuditEntry({
        method,
        target: redactUrl(validation.url),
        outcome: redirectedResponse.status >= 200 && redirectedResponse.status < 400 ? "accepted" : "failed",
        status: redirectedResponse.status,
        message:
          redirectedResponse.status >= 200 && redirectedResponse.status < 400
            ? "Safe redirect followed."
            : `Safe redirect returned HTTP ${redirectedResponse.status}.`,
        redirectTarget: redactUrl(resolvedLocation)
      });

      if (redirectedResponse.status >= 200 && redirectedResponse.status < 400) {
        response.json({
          ok: true,
          status: redirectedResponse.status,
          message: "The sender's unsubscribe redirect was opened successfully."
        });
        return;
      }

      response.status(502).json({
        ok: false,
        status: redirectedResponse.status,
        message: `The sender's unsubscribe redirect returned HTTP ${redirectedResponse.status}.`
      });
      return;
    }

    if (unsubscribeResponse.status >= 200 && unsubscribeResponse.status < 400) {
      addAuditEntry({
        method,
        target: redactUrl(validation.url),
        outcome: "accepted",
        status: unsubscribeResponse.status,
        message: method === "POST" ? "One-click unsubscribe accepted." : "Unsubscribe link opened."
      });
      response.json({
        ok: true,
        status: unsubscribeResponse.status,
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
      status: unsubscribeResponse.status,
      message: `Sender returned HTTP ${unsubscribeResponse.status}.`
    });
    response.status(502).json({
      ok: false,
      status: unsubscribeResponse.status,
      message: `The sender returned HTTP ${unsubscribeResponse.status}.`
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
