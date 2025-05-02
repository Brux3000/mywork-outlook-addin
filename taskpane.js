Office.onReady(() => {
  // Office is ready
});

function copySubject() {
  const item = Office.context.mailbox.item;
  const subject = item.subject;

  navigator.clipboard.writeText(subject).then(() => {
    document.getElementById('status').innerText = "Oggetto copiato!";
  }).catch(err => {
    document.getElementById('status').innerText = "Errore nella copia: " + err;
  });
}

