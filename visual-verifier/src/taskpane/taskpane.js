/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "./base64Image";

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // document.getElementById("run").onclick = run;
  }
  // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
  console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
  }


// Assign event handlers and other initialization logic.
  document.getElementById("insert-paragraph").onclick = insertParagraph;
  document.getElementById("insert-image").onclick = insertImage;
  document.getElementById("insert-html").onclick = insertHTML;
  document.getElementById("create-content-control").onclick = createContentControl;
  document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
});

function insertParagraph() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.body;
      
      docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                        "Start");
      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18
    });
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertImage() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert an image.
      context.document.body.insertInlinePictureFromBase64(base64Image, "End");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertHTML() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a string of HTML.
      var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function createContentControl() {
  Word.run(function (context) {

      // TODO1: Queue commands to create a content control.
      var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
serviceNameContentControl.title = "Service Name";
serviceNameContentControl.tag = "serviceName";
serviceNameContentControl.appearance = "Tags";
serviceNameContentControl.color = "blue";
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function replaceContentInControl() {
  Word.run(function (context) {

      // TODO1: Queue commands to replace the text in the Service Name
      //        content control.
      var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "black";

    await context.sync();
  });
}
