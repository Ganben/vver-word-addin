/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "./base64Image";
// import { MerkleTools } from "../../node_modules/merkle-tools/merkletools";
// import { MerkleTools } from "merkle-tools";
// const { QRCode }  = require('qrcode');
var QRious = require('qrious');
const { MerkleTree } = require('merkletreejs');
const SHA256 = require('crypto-js/sha256');

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    document.getElementById("run").onclick = run;
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

const sleep = (milliseconds) => {
  return new Promise(resolve => setTimeout(resolve, milliseconds))
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
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    const paras = context.document.body.paragraphs;
    // change the paragraph color to blue.
    context.load(paras, "text");
    await context.sync();

    // var merkleTools = new MerkleTools();
    var textList = [];
    paras.items.forEach(function (item, index) {
    //   // console.log(item, index);
    //   merkleTools.addLeaf(item.text, true);
    textList.push(item.text);
    });
    const leaves = textList.map(x => SHA256(x));
    const tree = new MerkleTree(leaves, SHA256);
    const root = tree.getRoot().toString('hex');
    // // paragraph.font.color = "black";
    // merkleTools.makeTree();
    // if (merkleTools.getTreeReadyState()) {
    //   var treeRoot = merkleTools.getMerkleRoot();
    // } else {
    //   await sleep(1000);
    //   var treeRoot = merkleTools.getMerkleRoot();
    // }
    // get first par
    
    var pf = context.document.body.paragraphs.getFirstOrNullObject();
    // pf.leftIndent = 80;
    // pf = pf.getNextOrNullObject();
    // pf.leftIndent = 90;
    // await context.sync();
    var i = 0;
    for (i = 0; i<textList.length; i++) {
      if (pf != null) {
        pf.leftIndent = 50;
        var qre = new QRious({
          value: textList[i],
          level: 'M',
          size: 100
        });
        var bimge = qre.toDataURL();
        pf.insertTable(1,1,"Before",[[textList[i]]]);
        pf.insertInlinePictureFromBase64(bimge.substr(22),"End");
        var qrs = new QRious({
          value: leaves[i].toString(),
          background:'blue',
          level: 'M',
          size:100
        });
        var bimgs = qrs.toDataURL();
        pf.insertInlinePictureFromBase64(bimgs.substr(22),"End");
        pf = pf.getNextOrNullObject();
        // await context.sync();
      } 
    }

    await context.sync();
    const para = context.document.body.insertParagraph("Paras:"+textList.length + " RootHash:0x"+root, Word.InsertLocation.end);
    
 
    para.font.color = "blue";
    para.leftIndent = 20;
    para.alignment = "Centered";
    // await context.sync();
    
    var qr = new QRious({
      value: root,
      background: 'red',
      backgroundAlpha: 0.8,
      level: 'M',
      // padding: 5,
      size: 200
    });
    var base64img = qr.toDataURL();
    const ppic = context.document.body.insertInlinePictureFromBase64(base64img.substr(22), "End");
    // const papic = context.document.body.insertParagraph(base64img.substr(23), 'End');
    ppic.alignment = "Centered";
    await context.sync();
    // do for each pars
    
    
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
}
