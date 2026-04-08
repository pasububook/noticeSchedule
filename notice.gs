function google_chat_webhook(message, webhookUrl) {
  const options = {
    "method": "post",
    "headers": {"Content-Type": "application/json; charset=UTF-8"},
    "payload": JSON.stringify({"text": message})
  };
  const response = UrlFetchApp.fetch(webhookUrl, options);
  console.log(response);
}
