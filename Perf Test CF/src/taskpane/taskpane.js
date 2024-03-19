/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  //document.getElementById("sideload-msg").style.display = "none";
  //document.getElementById("app-body").style.display = "flex";
  //document.getElementById("run").onclick = run;
  document.getElementById("calculate").onclick = calculate;
  document.getElementById("get_calculate").onclick = getcalculate;
  document.getElementById("persist").onclick = persist;
  document.getElementById("unpersist").onclick = unpersist;
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
      //range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function calculate() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.calculate();
      await context.sync();
      console.log("Calculated!");
    });
  } catch (error) {
    console.error(error);
  }
}

export async function getcalculate() {
  try {
    await Excel.run(async (context) => {
      var application = context.application;
      application.load("calculationMode, calculationState");
      await context.sync();
      var timestamp = new Date().toISOString();
      console.log(`[${timestamp}] calculationMode:`, application.calculationMode);
      console.log(`[${timestamp}] calculationState:`, application.calculationState);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function persist() {
  try {
    await Excel.run(async (context) => {
      Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
      Office.context.document.settings.saveAsync();
      console.log("Persisted!");
    });
  } catch (error) {
    console.error(error);
  }
}

export async function unpersist() {
  try {
    await Excel.run(async (context) => {
      Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
      Office.context.document.settings.saveAsync();
      console.log("Unpersisted!");
    });
  } catch (error) {
    console.error(error);
  }
}