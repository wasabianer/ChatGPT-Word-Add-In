/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  await Word.run(async (context) => {
    // Get the selected text range
    const range = context.document.getSelection();

    // Set the highlight color to red
    range.font.highlightColor = "red";

    // Synchronize the document state
    await context.sync();
  });
}
