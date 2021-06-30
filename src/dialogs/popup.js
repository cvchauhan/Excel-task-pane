(function () {
    "use strict";
    
    Office.onReady()
        .then(function() {
            document.getElementById("ok-button").onclick = sendStringToParentPage;           
        // TODO1: Assign handler to the OK button.

    });
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        alert(userName)
        Office.context.ui.messageParent(userName);
    }        
}());
