/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("openFormButton").onclick = openDataForm;
    document.getElementById("openFCFButton").onclick = openFCFForm;
  }
});

function openDataForm() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup-form.html",
    { height: 80, width: 50 },
    function (asyncResult) {
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
    }
  );
}

function openFCFForm() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup-form-fcf.html",
    { height: 80, width: 50 },
    function (asyncResult) {
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
    }
  );
}
function messageHandler(arg) {
  console.log(arg.message);
  // You can handle the message from the dialog box here
}
