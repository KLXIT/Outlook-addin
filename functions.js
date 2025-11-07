function checkWrikeId(event) {
  const item = Office.context.mailbox.item;
  const subject = item.subject;

  // Check for Wrike ID patterns
  const wrikeIdPatterns = [
    /\[?WRIKE-\d+\]?/i,
    /WID-\d+/i,
    /Task\s*#\d+/i,
    /\[?[A-Z]+-\d+\]?/i
  ];

  const hasWrikeId = wrikeIdPatterns.some(pattern => pattern.test(subject));

  if (hasWrikeId) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
      type: "informationalMessage",
      message: "✅ Wrike Task ID found in subject.",
      icon: "icon16",
      persistent: false
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(result.error.message);
      }
      event.completed();
    });
  } else {
    Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
      type: "informationalMessage",
      message: "⚠️ Please include a Wrike Task ID before sending.",
      icon: "icon16",
      persistent: false
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(result.error.message);
      }
      event.completed();
    });
  }
}
