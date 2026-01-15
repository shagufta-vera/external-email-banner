Office.onReady(() => {
  const item = Office.context.mailbox.item;
  if (!item || !item.from || !item.from.emailAddress) {
    return;
  }

  const sender = item.from.emailAddress;
  const internalDomains = ["15zn38.onmicrosoft.com"]; // CHANGE THIS

  const senderDomain = sender.split("@")[1]?.toLowerCase() || "";
  const isExternal = !internalDomains.includes(senderDomain);

  if (isExternal) {
    document.getElementById("sender").innerText = sender;
    document.body.style.display = "block";
  }
});
