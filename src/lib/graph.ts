import {
  createNestablePublicClientApplication,
  type IPublicClientApplication
} from "@azure/msal-browser";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.ReadWrite"];

export type GraphMessageSummary = {
  id: string;
  subject?: string;
  receivedDateTime?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  internetMessageHeaders?: { name: string; value: string }[];
};

let clientPromise: Promise<IPublicClientApplication> | null = null;

export function isGraphConfigured(clientId: string | null): clientId is string {
  return Boolean(clientId);
}

function getClient(clientId: string): Promise<IPublicClientApplication> {
  if (!clientPromise) {
    clientPromise = createNestablePublicClientApplication({
      auth: {
        clientId,
        authority: "https://login.microsoftonline.com/common"
      }
    });
  }
  return clientPromise;
}

export async function getGraphToken(clientId: string): Promise<string> {
  const client = await getClient(clientId);
  const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
  const request = { scopes: SCOPES, account: account ?? undefined };

  try {
    const result = await client.acquireTokenSilent(request);
    return result.accessToken;
  } catch {
    const result = await client.acquireTokenPopup(request);
    client.setActiveAccount(result.account);
    return result.accessToken;
  }
}

async function graphFetch<T>(clientId: string, path: string, init?: RequestInit): Promise<T> {
  const token = await getGraphToken(clientId);
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`Microsoft Graph request failed (HTTP ${response.status}). ${body.slice(0, 200)}`);
  }

  if (response.status === 204) {
    return undefined as T;
  }
  return (await response.json()) as T;
}

export function getCurrentMessageRestId(): string {
  const office = globalThis.Office;
  const item = office?.context?.mailbox?.item;

  if (!office || !item?.itemId) {
    throw new Error("Outlook did not expose the selected message id.");
  }

  return office.context.mailbox.convertToRestId(item.itemId, office.MailboxEnums.RestVersion.v2_0);
}

export async function moveMessageToDeletedItems(clientId: string, restMessageId: string): Promise<void> {
  await graphFetch(clientId, `/me/messages/${encodeURIComponent(restMessageId)}/move`, {
    method: "POST",
    body: JSON.stringify({ destinationId: "deleteditems" })
  });
}

export async function getParentFolderId(clientId: string, restMessageId: string): Promise<string> {
  const message = await graphFetch<{ parentFolderId: string }>(
    clientId,
    `/me/messages/${encodeURIComponent(restMessageId)}?$select=parentFolderId`
  );
  return message.parentFolderId;
}

export async function listFolderMessages(
  clientId: string,
  folderId: string,
  top = 50
): Promise<GraphMessageSummary[]> {
  const result = await graphFetch<{ value: GraphMessageSummary[] }>(
    clientId,
    `/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
      `?$top=${top}&$select=id,subject,from,receivedDateTime,internetMessageHeaders&$orderby=receivedDateTime desc`
  );
  return result.value ?? [];
}
