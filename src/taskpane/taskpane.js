  /*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("open-data-form").onclick = openDataForm;
    document.getElementById("openFormButton").onclick = openDataForm;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// function openDataForm() {
//   window.open('taskpane/popup-form.html', '_blank', 'width=800,height=600');
// }
function openDataForm() {
  Office.context.ui.displayDialogAsync('https://localhost:3000/popup-form.html',
    { height: 50, width: 50 },
    function (asyncResult) {
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
    });
}

function messageHandler(arg) {
  console.log(arg.message);
  // You can handle the message from the dialog box here
}
