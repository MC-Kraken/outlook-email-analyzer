/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run-sentiment-analysis").onclick = runSentimentAnalysis;
    document.getElementById("run-action-items").onclick = runActionItems;
    document.getElementById("run-questions").onclick = runQuestions;
    document.getElementById("formSubmission").onclick = submitForm;

    document.getElementById("currentEmail").innerHTML = "<b>Subject:</b> <br/>" + Office.context.mailbox.item.subject;
  }
});

const processTextUrl: string = "https://api.symbl.ai/v1/process/text";

async function runSentimentAnalysis() {
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", { asyncContext: "This is passed to the callback" }, function callback(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Process the result.value to extract the most recent email body
      formSentimentAnalysisPayload(result.value);
    } else {
      console.error("Failed to get email body:", result.error);
    }
  });
}

async function runActionItems() {
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", { asyncContext: "This is passed to the callback" }, function callback(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Process the result.value to extract the most recent email body
      formActionItemsPayload(result.value);
    } else {
      console.error("Failed to get email body:", result.error);
    }
  });
}

async function runQuestions() {
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", { asyncContext: "This is passed to the callback" }, function callback(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Process the result.value to extract the most recent email body
      formQuestionsPayload(result.value);
    } else {
      console.error("Failed to get email body:", result.error);
    }
  });
}

function formSentimentAnalysisPayload(emailText: string): void {
  const messages = [
    {
      payload: {
        content: emailText,
      },
    },
  ];

  processSentimentAnalysis(messages);
}

function formActionItemsPayload(emailText: string): void {
  const messages = [
    {
      payload: {
        content: emailText,
      },
    },
  ];

  processActionItems(messages);
}

function formQuestionsPayload(emailText: string): void {
  const messages = [
    {
      payload: {
        content: emailText,
      },
    },
  ];

  processQuestions(messages);
}

function processSentimentAnalysis(messages: any) {
  fetch(processTextUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer < token >",
    },
    body: JSON.stringify({ messages }),
  })
    .then((response) => response.json())
    .then((data) => {
      document.getElementById("sentiment-analysis").innerHTML = "<b>Processing...</b>";
      setTimeout((_) => {
        console.log("waiting for conversationId");
        console.log(data.conversationId);
        getTopicsWithSentiment(data.conversationId);
      }, 3000);
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

function processActionItems(messages: any) {
  fetch(processTextUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer < token >",
    },
    body: JSON.stringify({ messages }),
  })
    .then((response) => response.json())
    .then((data) => {
      document.getElementById("action-items").innerHTML = "<b>Processing...</b>";
      setTimeout((_) => {
        console.log("waiting for conversationId");
        console.log(data.conversationId);
        getActionItems(data.conversationId);
      }, 3000);
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

function processQuestions(messages: any) {
  fetch(processTextUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer < token >",
    },
    body: JSON.stringify({ messages }),
  })
    .then((response) => response.json())
    .then((data) => {
      document.getElementById("questions").innerHTML = "<b>Processing...</b>";
      setTimeout((_) => {
        console.log("waiting for conversationId");
        console.log(data.conversationId);
        getQuestions(data.conversationId);
      }, 3000);
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

function getTopicsWithSentiment(conversationId: string) {
  const secondUrl = `https://api.symbl.ai/v1/conversations/${conversationId}/topics?sentiment=true&parentRefs=false&customTopicVocabulary=`;
  fetch(secondUrl, {
    method: "GET",
    headers: {
      Authorization: "Bearer < token >",
    },
  })
    .then((response) => response.json())
    .then((data) => {
      const jsonString = JSON.stringify(data, null, 2); // The `null` and `2` arguments format the JSON with indentation for readability
      document.getElementById("sentiment-analysis").innerHTML =
        "<b>Sentiment Analysis:</b> <br/>" + `<pre>${jsonString}</pre>`;
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

function getActionItems(conversationId: string) {
  const getActionItemsUrl = `https://api.symbl.ai/v1/conversations/${conversationId}/action-items`;
  fetch(getActionItemsUrl, {
    method: "GET",
    headers: {
      Authorization: "Bearer < token >",
    },
  })
    .then((response) => response.json())
    .then((data) => {
      const jsonString = JSON.stringify(data, null, 2); // The `null` and `2` arguments format the JSON with indentation for readability
      document.getElementById("action-items").innerHTML = "<b>Action Items:</b> <br/>" + `<pre>${jsonString}</pre>`;
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

function getQuestions(conversationId: string) {
  const getQuestionsUrl = `https://api.symbl.ai/v1/conversations/${conversationId}/questions`;
  fetch(getQuestionsUrl, {
    method: "GET",
    headers: {
      Authorization: "Bearer < token >",
    },
  })
    .then((response) => response.json())
    .then((data) => {
      const jsonString = JSON.stringify(data, null, 2); // The `null` and `2` arguments format the JSON with indentation for readability
      document.getElementById("questions").innerHTML = "<b>Questions:</b> <br/>" + `<pre>${jsonString}</pre>`;
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("An error occurred while submitting the form.");
    });
}

export function submitForm(): void {
  const emailText: string = (document.getElementById("emailText") as HTMLTextAreaElement).value;

  const messages = [
    {
      payload: {
        content: emailText,
      },
    },
  ];

  processSentimentAnalysis(messages);
}
