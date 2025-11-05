Office.actions.associate("onMessageSendHandler", async function (event) {
  try {
    const item = Office.context.mailbox.item;

    const subjectResult = await new Promise((resolve) => {
      item.subject.getAsync((asyncResult) => {
        resolve(asyncResult.value || "");
      });
    });

    const subject = subjectResult.trim();
    const wrikePattern = /(WRK-\d{3,6}|#\d{3,6})/i;

    if (!wrikePattern.test(subject)) {
      const confirmed = await new Promise((resolve) => {
        Office.context.ui.displayDialogAsync(
          "https://raw.githubusercontent.com/KLXIT/Outlook-addin/main/wrike-reminder.html",
          { height: 30, width: 40 },
          (result) => {
            const dialog = result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
              dialog.close();
              resolve(arg.message === "yes");
            });
          }
        );
      });

      if (!confirmed) {
        event.completed({ allowEvent: false }); // cancel sending
        return;
      }
    }

    event.completed({ allowEvent: true });
  } catch (err) {
    console.error("Wrike Add-in error:", err);
    event.completed({ allowEvent: true });
  }
});
