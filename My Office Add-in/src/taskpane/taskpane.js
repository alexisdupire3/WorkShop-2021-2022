/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//import { exitCode } from "process";
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
    // await context.sync();

    // var list = [];

    // let firstParagraph = context.document.body.paragraphs.getFirst();
    // let currentParagraph = firstParagraph;
    // let lastParagraph = context.document.body.paragraphs.getLast();
    // let nombreParagraph = 1;

    // try {
    //   while (currentParagraph !== lastParagraph) {
    //     // if (currentParagraph == lastParagraph)
    //     //   return console.log("On a trouvé le dernier paragraphe, on abort")
    //     currentParagraph.font.color = "blue";
    //     nombreParagraph = nombreParagraph + 1;
    //     list.push(currentParagraph);
    //     console.log(`${list.length} tagada`);

    //     currentParagraph = await currentParagraph.getNext();
    //     await context.sync();
    //   }
    // } catch (e) {
    //   if (e.code == "ItemNotFound") {
    //     console.warn("[ERREUR] Pas possible d'aller plus loin dans la boucle")
    //   } else {
    //     console.warn("Erreur sur la boucle");
    //     console.warn(e);
    //     // currentParagraph.font.color = "red";
    //     return await context.sync();
    //   }
    // }

    // console.log("test")
    // console.log(`Il y'a ${list.length} paragraphes`)

    // console.log("Test")
    // console.log(list);

    const range = context.document.getSelection();
    
    const paragraphs = range.paragraphs;
    paragraphs.load();
    
    return context.sync()
      .then(() => {
        let ranges = paragraphs.items[0].getTextRanges(['.'], true);
        ranges.load();
        return context.sync()
          .then(() => {
            ranges.items.forEach((range) => {
              console.log(range.text);
            });
          });
      });


    // insert a paragraph at the end of the document.
    //const paragraph = context.document.body.paragraphs.getLast();

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";
    await context.sync();
  });
}
