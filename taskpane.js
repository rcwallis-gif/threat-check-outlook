let currentItemId = null;

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  if (!item) return;

  // Show subject in confirmation screen
  document.getElementById("subject-display").textContent = `"${item.subject}"`;
  currentItemId = item.itemId;

  // Auto-open confirmation when task pane loads
});

async function submitEmail() {
  document.getElementById("confirm-screen").classList.add("hidden");
  document.getElementById("loading-screen").classList.remove("hidden");

  const item = Office.context.mailbox.item;

  item.getAsFileAsync((asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      alert("Failed to get email: " + asyncResult.error.message);
      return;
    }

    const emlBase64 = asyncResult.value;

fetch('https://stamina-facelift-shell.ngrok-free.dev/analyze', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    emlBase64: emlBase64,
    subject: item.subject,
    itemId: currentItemId
  })
})
    .then(r => r.json())
    .then(data => {
      document.getElementById("loading-screen").classList.add("hidden");
      document.getElementById("report-screen").classList.remove("hidden");
      
      const reportDiv = document.getElementById("report-content");
      reportDiv.innerHTML = `
        <div class="flex items-center gap-2 mb-4">
          ${data.isThreat 
            ? `<span class="px-3 py-1 bg-red-100 text-red-700 text-xs font-bold rounded-full">THREAT DETECTED</span>` 
            : `<span class="px-3 py-1 bg-green-100 text-green-700 text-xs font-bold rounded-full">SAFE</span>`}
          <span class="font-mono text-xs text-gray-500">${data.confidence}</span>
        </div>
        <div class="text-gray-700 whitespace-pre-wrap">${data.report}</div>
      `;

      if (data.isThreat) {
        document.getElementById("delete-section").classList.remove("hidden");
      }
    })
    .catch(err => {
      alert("Backend error: " + err.message);
    });
  });
}

function cancel() {
  Office.context.ui.closeContainer();
}

function deleteEmail() {
  if (!currentItemId) return;
  
  const request = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                   xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <soap:Body>
        <m:DeleteItem>
          <m:ItemIds>
            <t:ItemId Id="${currentItemId}"/>
          </m:ItemIds>
          <m:DeleteType>MoveToDeletedItems</m:DeleteType>
        </m:DeleteItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("report-content").innerHTML += `<p class="text-green-600 font-bold mt-4">✅ Email moved to Deleted Items.</p>`;
      document.getElementById("delete-section").classList.add("hidden");
    } else {
      alert("Delete failed. You can delete it manually.");
    }
  });
}

function closeTaskpane() {
  Office.context.ui.closeContainer();
}