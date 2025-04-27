/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const storageItems = [
    "gptKey",
    "systemPrompt",
    "userPrompt",
    "filterPattern",
    "semaphoreCount",
    "gptUrl",
    "gptModel"
  ];

  const inputs = {};
  for (const item of storageItems) {
    // 使用正则表达式将驼峰命名转换为连字符分割的形式
    const localStorageKey = item.replace(/([A-Z])/g, "-$1").toLowerCase(); //   "glm-key-input": "glmKey",
    // 添加到 inputs 对象中，注意这里需要将属性名加上 '-input' 后缀
    inputs[localStorageKey + "-input"] = item;
    window[item] = getLocalStorage(item);
  }
  initializeInputs(inputs);
  initializeInputEvents(inputs);

  // 添加测试按钮的点击事件
  const testButton = document.getElementById("test-button");
  if (testButton) {
    testButton.addEventListener("click", handleTestButtonClick);
  }

  // 添加密码可见性切换功能
  const passwordToggle = document.getElementById("gpt-key-toggle");
  if (passwordToggle) {
    passwordToggle.addEventListener("click", togglePasswordVisibility);
  }

  // 添加复制示例代码功能
  const copyButton = document.getElementById("copy-example");
  if (copyButton) {
    copyButton.addEventListener("click", copyExampleCode);
  }
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
      element.value = getLocalStorage(inputs[inputId], element.value);
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

/**
 * 显示通知消息
 *
 * @param {string} message 消息内容
 * @param {string} type 消息类型: 'info', 'success', 'error', 'loading'
 * @param {number} duration 显示时长(毫秒)，默认5000ms，如果为0则不自动关闭
 * @returns {HTMLElement} 通知元素
 */
function showNotification(message, type = 'info', duration = 5000) {
  const container = document.getElementById('notification-container');

  // 创建通知元素
  const notification = document.createElement('div');
  notification.className = `notification notification-${type}`;
  notification.textContent = message;

  // 添加到容器
  container.appendChild(notification);

  // 如果设置了持续时间，则在指定时间后移除通知
  if (duration > 0) {
    setTimeout(() => {
      // 添加淡出动画
      notification.style.animation = 'fadeOut 0.3s ease-in-out forwards';

      // 动画结束后移除元素
      setTimeout(() => {
        if (container.contains(notification)) {
          container.removeChild(notification);
        }
      }, 300);
    }, duration);
  }

  return notification;
}

/**
 * 切换密码输入框的可见性
 */
function togglePasswordVisibility() {
  const passwordInput = document.getElementById("gpt-key-input");
  const eyeOpen = document.querySelector(".eye-open");
  const eyeClosed = document.querySelector(".eye-closed");

  if (passwordInput.type === "password") {
    // 显示密码
    passwordInput.type = "text";
    eyeOpen.style.display = "none";
    eyeClosed.style.display = "block";
  } else {
    // 隐藏密码
    passwordInput.type = "password";
    eyeOpen.style.display = "block";
    eyeClosed.style.display = "none";
  }
}

/**
 * 复制示例代码到剪贴板
 */
function copyExampleCode() {
  const codeText = document.querySelector('.example-code').textContent;

  // 使用 Clipboard API 复制文本
  navigator.clipboard.writeText(codeText)
    .then(() => {
      showNotification('复制成功！', 'success');
    })
    .catch(() => {
      showNotification('复制失败，请手动复制', 'error');
    });
}

/**
 * 处理测试按钮点击事件
 */
async function handleTestButtonClick() {
  try {
    // 显示测试开始的消息
    console.log("测试开始");

    // 检查API密钥是否已设置
    if (!window.gptKey) {
      showNotification("请先设置GPT密钥", "error");
      return;
    }

    // 创建一个简单的测试提示
    const testPrompt = "这是一个测试消息，请回复'测试成功'";

    // 构建请求消息
    const messages = [];
    if (window.systemPrompt && window.systemPrompt.trim() !== "") {
      messages.push({
        role: "system",
        content: window.systemPrompt
      });
    }

    messages.push({
      role: "user",
      content: testPrompt
    });

    // 显示正在测试的消息
    const loadingNotification = showNotification("正在测试API连接，请稍候...", "loading", 0);

    // 发送API请求
    const model = window.gptModel || "gpt-4-turbo";
    const url = window.gptUrl || "https://api.openai.com/v1/chat/completions";
    const apiKey = window.gptKey;

    // 确保API密钥不包含非ASCII字符
    const sanitizedApiKey = apiKey.replace(/[^\x00-\x7F]/g, "");

    // 构建请求头，确保所有值都是有效的ASCII字符
    const headers = new Headers();
    headers.append("Content-Type", "application/json");
    headers.append("Authorization", "Bearer " + sanitizedApiKey);

    const response = await fetch(url, {
      method: "POST",
      headers: headers,
      body: JSON.stringify({
        model: model,
        messages: messages,
      }),
    });

    // 移除加载通知
    const container = document.getElementById('notification-container');
    if (container.contains(loadingNotification)) {
      container.removeChild(loadingNotification);
    }

    const json = await response.json();
    if (response.status !== 200 && json.error) {
      throw new Error(json.error.message);
    }

    const result = json.choices[0].message.content;

    // 显示测试结果
    showNotification("测试结果: " + result, "success");
    console.log("测试结果:", result);

  } catch (error) {
    // 显示错误信息
    showNotification("测试失败: " + error.message, "error");
    console.error("测试失败:", error);
  }
}
