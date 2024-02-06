/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const storageItems = ["glmKey", "gptKey", "sparkKey", "deepseekKey", "systemPrompt"];
  
  const inputs = {};
  for (const item of storageItems) {
    // 使用正则表达式将驼峰命名转换为连字符分割的形式
    const localStorageKey = item.replace(/([A-Z])/g, "-$1").toLowerCase();  //   "glm-key-input": "glmKey",
    // 添加到 inputs 对象中，注意这里需要将属性名加上 '-input' 后缀
    inputs[localStorageKey + "-input"] = item;
    window[item] = getLocalStorage(item);
  }
  initializeInputs(inputs);
  initializeInputEvents(inputs);
  console.log("taskpanel onReady");
});

/**
 * 运行Excel操作
 * document.getElementById("run").onclick = run;
 *
 * @returns {Promise<void>}
 */
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

/**
 * 获取本地存储中的指定键值对应的值。
 *
 * @param key 存储的键名。
 * @param defaultValue 默认值，当键值不存在时返回该值。默认为空字符串。
 * @returns 返回键值对应的值，如果键值不存在则返回默认值。
 */
function setLocalStorage(key, value) {
  localStorage.setItem(key, value);
}

/**
 * 获取本地存储中的值
 *
 * @param key 键名
 * @param defaultValue 默认值，默认为空字符串
 * @returns 返回键名对应的值，若不存在则返回默认值
 */
function getLocalStorage(key, defaultValue = "") {
  return localStorage.getItem(key) || defaultValue;
}

/**
 * 初始化输入框
 *
 * @param inputs 包含输入框id和值的对象
 */
function initializeInputs(inputs) {
  Object.keys(inputs).forEach((inputId) => {
    const element = document.getElementById(inputId);
    if (element) {
      element.value = getLocalStorage(inputs[inputId]);
    }
  });
}
/**
 * 初始化输入框事件
 *
 * @param inputs 输入框对象，键为输入框id，值为输入框值
 */
function initializeInputEvents(inputs) {
  const container = document.getElementById("input-container"); // 假设所有输入框都在一个容器内
  container.addEventListener("change", (event) => {
    const inputId = event.target.id;
    // 检查event.target是否是inputs对象中的有效输入框
    if (inputs.hasOwnProperty(inputId)) {
      const localStorageKey = inputs[inputId];
      window[localStorageKey] = event.target.value;
      setLocalStorage(localStorageKey, event.target.value);
    }
  });
}
