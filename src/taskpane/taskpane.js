/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.OneNote) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        //page.title = document.getElementById("project_title").value

        // Queue a command to add an outline to the page.
        var html = "Startas: <br>"
        // https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.outline?view=onenote-js-1.1
        let values = [  ['Užduotis', 'Startas'], ['Uzregistruoti','2019-11-10'] ]
        page.addOutline(40, 80, html).appendTable(2, 2, values).appendColumn(["Trukmė"]);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
}
