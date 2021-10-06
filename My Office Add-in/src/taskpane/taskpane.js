/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//import { lineBreak } from "acorn";

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
    let paragraph = context.document.body.paragraphs.getFirst();
    paragraph.font.color = "blue";
    while (paragraph !== null) {
      paragraph.font.color = "blue";
      paragraph = paragraph.getNextOrNullObject();
    }

    // insert a paragraph at the end of the document.
    //const paragraph = context.document.body.paragraphs.getLast();

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";

    await context.sync();
  });
}
