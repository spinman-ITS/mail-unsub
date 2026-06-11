export type MarketoFormTarget = {
  submitUrl: string;
  formId: string;
  munchkinId: string;
};

// Marketo preference centers embed a formDescriptor JSON blob. The standard
// system field "Unsubscribed" is the "unsubscribe from all" checkbox.
export function findMarketoUnsubscribeForm(html: string, pageUrl: string): MarketoFormTarget | null {
  if (!html.includes('"Name":"Unsubscribed"')) {
    return null;
  }

  const formId = html.match(/"Id":(\d+),"Vid"/)?.[1];
  const munchkinId = html.match(/"munchkinId":"([^"]+)"/)?.[1];
  const actionRaw = html.match(/"action":"([^"]+)"/)?.[1];

  if (!formId || !munchkinId || !actionRaw) {
    return null;
  }

  const action = actionRaw.replace(/\\\//g, "/");

  try {
    return {
      submitUrl: new URL(action, pageUrl).toString(),
      formId,
      munchkinId
    };
  } catch {
    return null;
  }
}

export function buildMarketoUnsubscribeBody(email: string, target: MarketoFormTarget): string {
  return new URLSearchParams({
    formid: target.formId,
    munchkinId: target.munchkinId,
    Email: email,
    Unsubscribed: "yes"
  }).toString();
}
