/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("delete-footer-button").onclick = deleteAllFooters;
    }
});

// Function to delete all footers in the active document
async function deleteAllFooters() {
    await Word.run(async (context) => {
        try {
            const sections = context.document.sections;
            sections.load("body/style")
            await context.sync();

            sections.items.forEach(function (section) {
                var footer = section.getFooter(Word.HeaderFooterType.primary)
                footer.clear()
            });
            await context.sync();

            showNotification("All footers have been deleted.");
        } catch (e) {
            errorHandler(e);
        }
    })
}

//$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
function errorHandler(error) {
    var message = "Error: " + error;
    showNotification(message);
    console.log(message);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

// Helper function for displaying notifications
function showNotification(text) {
    document.getElementById("notification-text").innerHTML = text;
}