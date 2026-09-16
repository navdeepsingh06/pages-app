Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("uploadButton").onclick = uploadEmailToDMS;
    }
});

function uploadEmailToDMS() {
    const item = Office.context.mailbox.item;
    
    // Get the REST ID or EWS ID to fetch the full MIME content on your backend,
    // or use getFileAsync for attachments/eml conversion.
    item.getStartupDataAsync ? console.log("Initializing upload...") : null;

    // Example using mailbox REST API token to get item content
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const accessToken = asyncResult.value;
            const itemId = item.itemId;
            const restUrl = Office.context.mailbox.restUrl;
            
            // Construct the endpoint to fetch the MIME/EML message
            const getMessageUrl = `${restUrl}/v2.0/me/messages/${itemId}/$value`;

            // Send ID and token to your backend service which handles the DMS upload
            sendToBackendDMS(getMessageUrl, accessToken);
        } else {
            console.error("Failed to get token: " + asyncResult.error.message);
        }
    });
}

function sendToBackendDMS(url, token) {
    fetch('https://your-dms-backend-api.com', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({ messageUrl: url })
    })
    .then(response => response.json())
    .then(data => {
        console.log("Successfully uploaded to DMS", data);
    })
    .catch(error => {
        console.error("Error uploading email:", error);
    });
}
