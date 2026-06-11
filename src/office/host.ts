export function hasOutlookMessageContext(): boolean {
  return Boolean(globalThis.Office?.context?.mailbox?.item?.getAllInternetHeadersAsync);
}

export async function waitForOutlookMessageContext(timeoutMs = 10000): Promise<boolean> {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    if (globalThis.Office?.onReady) {
      await globalThis.Office.onReady();
    }

    if (hasOutlookMessageContext()) {
      return true;
    }

    await delay(100);
  }

  return false;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}
