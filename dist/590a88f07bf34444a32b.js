function openDataForm(){Office.context.ui.displayDialogAsync("https://localhost:3000/popup-form.html",{height:80,width:50},(function(e){e.value.addEventHandler(Office.EventType.DialogMessageReceived,messageHandler)}))}function openFCFForm(){Office.context.ui.displayDialogAsync("https://localhost:3000/popup-form-fcf.html",{height:80,width:50},(function(e){e.value.addEventHandler(Office.EventType.DialogMessageReceived,messageHandler)}))}function messageHandler(e){console.log(e.message)}Office.onReady((function(e){e.host===Office.HostType.Excel&&(document.getElementById("openFormButton").onclick=openDataForm,document.getElementById("openFCFbutton").onclick=openFCFForm)}));