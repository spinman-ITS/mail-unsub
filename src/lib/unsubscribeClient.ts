export type OneClickResponse = {
  ok: boolean;
  message: string;
  status?: number;
  requiresBrowser?: boolean;
  finalUrl?: string;
};

export type AppConfig = {
  aadClientId: string | null;
};

export type UnsubscribedSender = {
  userEmail: string;
  senderAddress: string;
  senderDomain: string;
  method: string;
  createdAt: string;
};

const apiBaseUrl = import.meta.env.DEV ? "https://localhost:8787" : "";

export async function performOneClickUnsubscribe(url: string): Promise<OneClickResponse> {
  return postUnsubscribeRequest("/api/unsubscribe", url);
}

export async function performBodyLinkUnsubscribe(url: string): Promise<OneClickResponse> {
  return postUnsubscribeRequest("/api/unsubscribe-link", url);
}

export async function fetchAppConfig(): Promise<AppConfig> {
  try {
    const response = await fetch(`${apiBaseUrl}/api/config`);
    if (!response.ok) {
      return { aadClientId: null };
    }
    return (await response.json()) as AppConfig;
  } catch {
    return { aadClientId: null };
  }
}

export async function recordUnsubscribedSender(
  userEmail: string,
  senderAddress: string,
  method: string
): Promise<void> {
  await fetch(`${apiBaseUrl}/api/unsubscribed`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ userEmail, senderAddress, method })
  }).catch(() => undefined);
}

export async function fetchUnsubscribedSenders(userEmail: string): Promise<UnsubscribedSender[]> {
  try {
    const response = await fetch(`${apiBaseUrl}/api/unsubscribed?userEmail=${encodeURIComponent(userEmail)}`);
    if (!response.ok) {
      return [];
    }
    const payload = (await response.json()) as { senders?: UnsubscribedSender[] };
    return payload.senders ?? [];
  } catch {
    return [];
  }
}

async function postUnsubscribeRequest(path: string, url: string): Promise<OneClickResponse> {
  const response = await fetch(`${apiBaseUrl}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ url })
  });

  const payload = (await response.json().catch(() => null)) as Partial<OneClickResponse> | null;

  if (!response.ok) {
    return {
      ok: false,
      message: payload?.message || "The unsubscribe request could not be completed.",
      status: response.status
    };
  }

  return {
    ok: Boolean(payload?.ok),
    message: payload?.message || "The unsubscribe request was sent.",
    status: payload?.status,
    requiresBrowser: payload?.requiresBrowser,
    finalUrl: payload?.finalUrl
  };
}
