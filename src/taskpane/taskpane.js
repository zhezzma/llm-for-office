/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  //document.getElementById("run").onclick = run;
  document.getElementById("glm-key-input").value = localStorage.getItem("glm-key");
  document.getElementById("gpt-key-input").value = localStorage.getItem("gpt-key");
  document.getElementById("glm-key-input").onchange = storeGLMValue;
  document.getElementById("gpt-key-input").onchange = storeGPTValue;
  console.log("taskpanel onReady")
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

export function storeGLMValue() {
  window.glmKey = document.getElementById('glm-key-input').value;
  localStorage.setItem('glm-key', window.glmKey);
}


export function storeGPTValue() {
  window.gptKey = document.getElementById('gpt-key-input').value;
  localStorage.setItem('gpt-key', window.gptKey);
}

