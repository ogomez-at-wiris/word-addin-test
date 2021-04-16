/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    document.getElementById("alertSpeaker").innerHTML =
      "Welcome to MathType. To start working, please use the show task pane option in the home ribbon.";
    console.log("hi");
    const paragraph = context.document.body.insertParagraph("This is MathType", Word.InsertLocation.end);
    paragraph.font.color = "blue";
    await context.sync();
  });
}
