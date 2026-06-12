// Open a URL in an Office.js dialog. Resolves when the user closes the
// dialog (or it fails to open). Uses displayDialogAsync so it works on all
// Outlook platforms — window.open is blocked inside the taskpane iframe.
export function openDialogAndWait(url: string): Promise<void> {
  return new Promise((resolve) => {
    const office = globalThis.Office;
    if (!office?.context?.ui?.displayDialogAsync) {
      window.open(url, "_blank", "noopener,noreferrer");
      resolve();
      return;
    }

    office.context.ui.displayDialogAsync(
      url,
      { height: 70, width: 50, promptBeforeOpen: false },
      (result) => {
        if (result.status !== office.AsyncResultStatus.Succeeded || !result.value) {
          resolve();
          return;
        }

        const dialog = result.value;
        const done = () => {
          try {
            dialog.close();
          } catch {
            // Already closed.
          }
          resolve();
        };

        dialog.addEventHandler(office.EventType.DialogEventReceived, done);
        // 12 = userClosed; any DialogMessageReceived also signals the user is done.
        dialog.addEventHandler(office.EventType.DialogMessageReceived, done);
      }
    );
  });
}
