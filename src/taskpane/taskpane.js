/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("summary").onclick = summary;
  }
});

export async function run() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;

  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext:"This is passed to the callback" },
    function callback(result) {
        if (result.status === "succeeded") {
            console.log("The body is " + result.value);
            document.getElementById("item-content").innerHTML = result.value;

        } else {
            console.log("Failed to get body. Error: " + result.error);
        }
    }
);
  }
export async function summary() {
           const content = document.getElementById("item-content").innerHTML;
           const url = "https://<LLMAPI>.azurewebsites.net/api/summarize";
           
            const response = await fetch(url, {
              method: "POST",
              headers: {
                "Content-Type": "text/plain",
              },
              body: content,
            }).then((response) => response.text())
            .then((data) =>
            {
              console.log(data);
              document.getElementById("item-summarize").innerHTML = data;
            } );
          }


