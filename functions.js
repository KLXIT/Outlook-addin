function checkWrikeId(event) {
  Office.context.mailbox.item.subject.getAsync((res) => {
    if (!res.value || !res.value.match(/\bWrike\b/i)) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
        type: "informationalMessage",
        message: "⚠️ Please include a Wrike Task ID in the subject before sending.",
        icon: "icon16",
        persistent: false
      });
    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("wrikeCheck", {
        type: "informationalMessage",
        message: "✅ Wrike Task ID detected.",
        icon: "icon16",
        persistent: false
      });
    }
  });
  event.completed();
}
