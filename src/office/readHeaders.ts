export function readCurrentMessageHeaders(): Promise<string> {
  return new Promise((resolve, reject) => {
    const office = globalThis.Office;

    if (!office?.context?.mailbox?.item?.getAllInternetHeadersAsync) {
      reject(new Error("Open this add-in from a message in Outlook to read unsubscribe headers."));
      return;
    }

    office.context.mailbox.item.getAllInternetHeadersAsync((result) => {
      if (result.status === office.AsyncResultStatus.Succeeded) {
        resolve(result.value ?? "");
        return;
      }

      reject(new Error(result.error?.message || "Outlook could not read this message's internet headers."));
    });
  });
}

export function readCurrentMessageBodyHtml(): Promise<string> {
  return new Promise((resolve, reject) => {
    const office = globalThis.Office;
    const body = office?.context?.mailbox?.item?.body;

    if (!office || !body?.getAsync) {
      reject(new Error("Outlook did not expose the selected message body."));
      return;
    }

    body.getAsync(office.CoercionType.Html, (result) => {
      if (result.status === office.AsyncResultStatus.Succeeded) {
        resolve(result.value ?? "");
        return;
      }

      reject(new Error(result.error?.message || "Outlook could not read this message's body."));
    });
  });
}
