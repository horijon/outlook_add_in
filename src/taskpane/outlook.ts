/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getApiEndpoints } from './config';

function showStatus(message: string): void {
  const statusContainer = document.getElementById("status-container");
  const statusMessage = document.getElementById("status-message");

  if (statusContainer && statusMessage) {
    statusMessage.textContent = message;
    statusContainer.style.display = "block";
  }
}

function hideStatus(): void {
  const statusContainer = document.getElementById("status-container");
  if (statusContainer) {
    statusContainer.style.display = "none";
  }
}

function showError(message: string): void {
  const errorContainer = document.getElementById("error-container");
  const errorMessage = document.getElementById("error-message");

  if (errorContainer && errorMessage) {
    errorMessage.textContent = message;
    errorContainer.style.display = "block";
  }
}

function hideError(): void {
  const errorContainer = document.getElementById("error-container");
  if (errorContainer) {
    errorContainer.style.display = "none";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate-response").onclick = generateResponse;
    document.getElementById("contact-support").onclick = contactSupport;
  }
});

// Function to convert markdown to HTML similar to the Apps Script version
function markdownToHtml(markdown: string): string {
  if (!markdown) return '';

  return markdown
    .replace(/(#{1,6})\s*(.*)/g, (match, hashes, text) =>
      `<h${hashes.length}>${text}</h${hashes.length}>`)
    .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
    .replace(/\*(.*?)\*/g, '<i>$1</i>')
    .replace(/`(.*?)`/g, '<code>$1</code>')
    .replace(/\[(.*?)\]\((.*?)\)/g, '<a href="$2">$1</a>')
    .replace(/\n\n/g, '</p><p>')
    .replace(/\n/g, '<br>');
}

function contactSupport(): void {
  window.open("https://govstream.ai", "_blank");
}

// Get email body as plain text
function getEmailBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error("Failed to get email body: " + result.error.message));
        }
      }
    );
  });
}

export async function generateResponse() {
  try {
    hideError();
    showStatus("Generating response...");

    const item = Office.context.mailbox.item;
    if (!item) {
      throw new Error("No email message found");
    }

    // Get email details
    const subject = item.subject || "";
    const emailBody = await getEmailBody();
    const messageId = item.internetMessageId || item.itemId || "";
    const receivedTime = item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString();

    // In a real app, you might want to get an auth token here
    // const token = await getAuthToken();

    // Get API endpoints from config
    const apiEndpoints = getApiEndpoints();

    const response = await fetch(apiEndpoints.emailProcess, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        // Add authentication if needed 
        // "Authorization": "Bearer " + token,
      },
      body: JSON.stringify({
        subject: subject,
        emailBody: emailBody,
        messageId: messageId,
        receivedTime: receivedTime
      }),
    });

    const data = await response.json();

    if (data?.status !== "success") {
      throw new Error(data.message || "An unknown error occurred");
    }

    // Convert the response from markdown to HTML
    const htmlContent = markdownToHtml(data.response);

    // Create metadata similar to the Apps Script version
    const metadata = `
     <div style="display:none;">
        <p>Original Message ID: ${messageId}</p>
        <p>Subject: ${subject}</p>
      </div>
    `;

    // Create a draft reply with the generated content
    item.displayReplyForm({
      htmlBody: htmlContent + metadata,
    });

    hideStatus();
  } catch (error) {
    console.error("Error generating response:", error);
    hideStatus();
    showError(error.message || "An unknown error occurred");
  }
}
