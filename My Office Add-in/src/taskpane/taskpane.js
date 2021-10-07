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

    const range = context.document.body;
    const paragraphs = range.paragraphs;
    paragraphs.load();

    return await context.sync()
      .then(async () => {
        //Tableau qui stock les paragraphes
        // /** @type {Word.ParagraphCollection} */
        // let arrayRanges = [];
        // let i = 0;

        //Boucle while qui recup tous les paragraphes en découpant à chaque "."
        // while (paragraphs.items.length > i) {
        //   console.log(paragraphs.items[i]);
        //   arrayRanges.push(paragraphs.items[i].getTextRanges(['.'], true));
        //   i++;
        // }
        //paragraphs.items[0].getTextRanges(['.'], true);

        let arrayRanges = [];
        let ranges;
        for (var i = 0; i < paragraphs.items.length; i++) {
          ranges = paragraphs.items[i].getTextRanges(['.'], true);
          ranges.load();
          context.sync()
            .then(() => {
              ranges.items.forEach((range) => {
                //console.log(range);
              });
            });
        }
        console.log(ranges.getFirst());
        //  console.log(arrayRanges)
        //console.log(paragraphs.items[0]);
        // console.log(paragraphs.items.length);
        // ranges.load();

        // return context.sync()
        //   .then(() => {
        //     arrayRanges.forEach((range) => {
        //       console.log(range.text);
        //     });
        //   });

      });
    await context.sync();
  });
}