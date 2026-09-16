"use strict";

let isInitialized = false;

const emailData = {
  subject: "",
  sender: "",
  receivedDate: "",
  internetMessageId: "",
  bodyPreview: ""
};

Office.onReady((info) => {
  if (isInitialized) return;
  isInitialized = true;

  if (info.host !== Office.HostType.Outlook) {
    showLoadError("This task pane must be opened from an Outlook email.");
    return;
  }

  document.getElementById("claim-form").addEventListener("submit", submitClaim);
  loadSelectedEmail();
});

function loadSelectedEmail() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showLoadError("No email is selected. Select an email and open the task pane again.");
    return;
  }

  emailData.subject = item.subject || "(No subject)";
  emailData.sender = item.from && item.from.emailAddress ? item.from.emailAddress : "Not available";
  emailData.receivedDate = formatDate(item.dateTimeCreated);
  emailData.internetMessageId = item.internetMessageId || "";

  renderEmailSummary();
  item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.bodyPreview = result.value.replace(/\s+/g, " ").trim().slice(0, 500);
    }
    setLoadStatus("Selected email loaded.");
    document.getElementById("submit-button").disabled = false;
  });
}

function renderEmailSummary() {
  document.getElementById("email-subject").textContent = emailData.subject;
  document.getElementById("email-sender").textContent = emailData.sender;
  document.getElementById("email-received-date").textContent = emailData.receivedDate;
}

function submitClaim(event) {
  console.log("Submit button clicked.");
  event.preventDefault();
  const claimNumber = document.getElementById("claim-number").value.trim();
  const errorElement = document.getElementById("claim-number-error");

  errorElement.textContent = "";
  if (!claimNumber) {
    errorElement.textContent = "Claim Number is required.";
    document.getElementById("claim-number").focus();
    return;
  }

  const payload = { claimNumber, ...emailData };
  console.log("ClaimCenter mock payload:", payload);

  document.getElementById("result-message").textContent = "The payload was prepared and logged to the browser console.";
  document.getElementById("payload-output").textContent = JSON.stringify(payload, null, 2);
  document.getElementById("result").hidden = false;
  uploadEmailFile(claimNumber);
}

function formatDate(date) {
  if (!date) {
    return "Not available";
  }
  return new Date(date).toLocaleString();
}

function setLoadStatus(message) {
  const status = document.getElementById("load-status");
  status.textContent = message;
  status.classList.remove("error");
}

function showLoadError(message) {
  const status = document.getElementById("load-status");
  status.textContent = message;
  status.classList.add("error");
  document.getElementById("submit-button").disabled = true;
}

function setStatus(message) {
    document.getElementById("result-message").innerText = message;
}

function sanitizeFileName(value) {
    return value.replace(/[\\/:*?"<>|]/g, "_");
}

async function uploadEmailFile(claimNumber) {
  try {
    setStatus("Acquiring access token…");
    console.log("[Upload] Acquiring access token before upload.");

    /*let accessToken;
    try {
      accessToken = await getAccessToken();
      console.log("[Upload] Access token ready.");
    } catch (tokenErr) {
      console.error("[Upload] Token acquisition failed:", tokenErr);
      setStatus("Authentication failed: " + tokenErr.message);
      return;
    }*/

    const item = Office.context.mailbox.item;

    setStatus("Exporting email…");
    console.log("[Upload] Calling item.getAsFileAsync().");

    item.getAsFileAsync(async (result) => {
      try {
        console.log("[Upload] getAsFileAsync result status:", result.status);

        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          throw new Error("Failed to retrieve email as file.");
        }

        const base64Content = result.value;
        const fileName = sanitizeFileName(item.subject || "email") + ".eml";

        console.log("[Upload] Email exported successfully.");
        console.log("[Upload] File name:", fileName);
        console.log("[Upload] Base64 content length:", base64Content.length);

        const payload = {
          claimNumber,
          fileName,
          contentType: "message/rfc822",
          base64Content
        };

        console.log("[Upload] Sending payload to /api/upload:", {
          claimNumber: payload.claimNumber,
          fileName: payload.fileName,
          contentType: payload.contentType,
          base64Length: payload.base64Content.length
        });

        setStatus("Uploading to Spring Boot…");

        const response = await fetch(`/api/upload`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": "Basic " + btoa("su:gw")
          },
          body: JSON.stringify(payload)
        });

        console.log("[Upload] Response status:", response.status, response.statusText);

        if (!response.ok) {
          const errorBody = await response.text();
          console.error("[Upload] Error response body:", errorBody);
          throw new Error(
            "Server HTTP " + response.status + " — " + errorBody
          );
        }

        const responseText = await response.text();
        console.log("[Upload] Spring Boot response:", responseText);

        setStatus(
          "Success!\n\nClaim Number: " + claimNumber + "\n\nFile Name: " + fileName
        );
        document.getElementById("result").hidden = false;

      } catch (err) {
        console.error("[Upload] Error inside getAsFileAsync callback:", err);
        setStatus(err.message);
      }
    });

  } catch (err) {
    console.error("[Upload] Unexpected error:", err);
    setStatus(err.message);
  }
}
