/* global clearnumbererval, console, setnumbererval */
//https://learn.microsoft.com/zh-cn/office/dev/add-ins/excel/

const CryptoJS = require("crypto-js");
/**
 * 使用chatgpt生成你想要的数据
 * @customfunction GPT
 * @param {string} prompt 输入提示
 * @param {string} [value] 单元格的内容
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function gpt(prompt, value, fillOffset, invocation) {
  if (!window.gptKey) {
    return "apiKey 未设置";
  }

  const validateResult = validateUserPrompt(prompt, value);
  if (validateResult.error) {
    return validateResult.errorMsg;
  }

  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;

  let result = "";

  try {
    const model = "gpt-3.5-turbo";
    const url = "https://openai-forward-s4pz.onrender.com/v1/chat/completions";
    const apiKey = window.gptKey;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model: model,
        messages: validateResult.messages,
      }),
    });
    const json = await response.json();
    if (response.status !== 200 && json.error) {
      throw new Error(json.error.message);
    }
    result = json.choices[0].message.content;
  } catch (error) {
    result = error.message;
  }

  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
}

/**
 * 使用DEEPSEEK生成你想要的数据
 * @customfunction DEEPSEEK
 * @param {string} prompt 输入提示
 * @param {string} [value] 单元格的内容
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function deepseek(prompt, value, fillOffset, invocation) {
  if (!window.deepseekKey) {
    return "apiKey 未设置";
  }

  const validateResult = validateUserPrompt(prompt, value);
  if (validateResult.error) {
    return validateResult.errorMsg;
  }

  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;

  let result = "";

  try {
    const model = "deepseek-chat";
    const url = "https://api.deepseek.com/v1/chat/completions";
    const apiKey = window.deepseekKey;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model: model,
        messages: validateResult.messages,
      }),
    });
    const json = await response.json();
    if (response.status !== 200 && json.detail) {
      throw new Error(json.detail);
    }
    result = json.choices[0].message.content;
  } catch (error) {
    result = error.message;
  }

  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
}

/**
 * 使用chatglm生成你想要的数据
 * @customfunction GLM
 * @param {string} prompt 输入提示
 * @param {string} [value] 单元格的内容
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function glm(prompt, value, fillOffset, invocation) {
  if (!window.glmKey) {
    return "apiKey 未设置";
  }

  const validateResult = validateUserPrompt(prompt, value);
  if (validateResult.error) {
    return validateResult.errorMsg;
  }

  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;
  let result = "";
  try {
    const url = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
    const apiKey = generateGLMToken(window.glmKey, 3600);
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model: "glm-4",
        messages: validateResult.messages,
      }),
    });
    const json = await response.json();
    if (response.status !== 200 && json.error) {
      throw new Error(json.error.message);
    }
    result = json.choices[0].message.content;
  } catch (error) {
    result = error.message;
  }
  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
}

/**
 * 生成GLM Token
 *
 * @param apikey API密钥，格式为"id.secret"
 * @param expSeconds 过期时间（秒）
 * @returns 返回GLM Token字符串
 */
function generateGLMToken(apikey, expSeconds) {
  const [id, secret] = apikey.split(".");
  const payload = {
    api_key: id,
    exp: Math.floor(Date.now() / 1000) + expSeconds,
    timestamp: Math.floor(Date.now() / 1000),
  };
  const header = {
    alg: "HS256",
    sign_type: "SIGN",
  };
  const stringifiedHeader = CryptoJS.enc.Utf8.parse(JSON.stringify(header));
  const stringifiedPayload = CryptoJS.enc.Utf8.parse(JSON.stringify(payload));
  const encodedHeader = CryptoJS.enc.Base64.stringify(stringifiedHeader);
  const encodedPayload = CryptoJS.enc.Base64.stringify(stringifiedPayload);
  const signature = CryptoJS.HmacSHA256(encodedHeader + "." + encodedPayload, secret);
  const encodedSignature = CryptoJS.enc.Base64.stringify(signature);
  return encodedHeader + "." + encodedPayload + "." + encodedSignature;
}

/**
 * 使用星火生成你想要的数据
 * @customfunction SPARK
 * @param {string} prompt 输入提示
 * @param {string} [value] 单元格的内容
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function spark(prompt, value, fillOffset, invocation) {
  if (!window.sparkKey) {
    return "apiKey 未设置";
  }

  const validateResult = validateUserPrompt(prompt, value);
  if (validateResult.error) {
    return validateResult.errorMsg;
  }
  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;
  let result = "";
  const version = "v3.5";
  const domain = `general${version}`;
  const [APPID, APISecret, APIKey] = window.sparkKey.split(".");
  const url = getSparkUrl(APISecret, APIKey, version);
  const ttsWS = new WebSocket(url);
  try {
    // 等待连接打开
    await new Promise((resolve) => (ttsWS.onopen = resolve));

    // 发送消息
    ttsWS.send(
      JSON.stringify({
        header: {
          app_id: APPID,
          uid: "godgodgame",
        },
        parameter: {
          chat: {
            domain: domain,
            temperature: 0.5,
            max_tokens: 1024,
          },
        },
        payload: {
          message: {
            text: validateResult.messages,
          },
        },
      })
    );

    // 接收消息
    while (true) {
      const e = await new Promise((resolve) => (ttsWS.onmessage = resolve));
      const jsonData = JSON.parse(e.data);
      // 提问失败
      if (jsonData.header.code !== 0) {
        throw new Error(jsonData.header.message);
      }
      result += jsonData.payload.choices.text[0].content;
      // 接收完成
      if (jsonData.header.code === 0 && jsonData.header.status === 2) {
        break;
      }
    }
  } catch (error) {
    result = error.message;
  } finally {
    // 清理事件监听器
    ttsWS.onopen = null;
    ttsWS.onmessage = null;
    ttsWS.onclose = null;
    ttsWS.close();
  }
  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
}

/**
 * 获取Spark API的WebSocket URL
 *
 * @param apiSecret API密钥
 * @param apiKey API密钥
 * @param version API版本号
 * @returns 返回WebSocket URL
 */
function getSparkUrl(apiSecret, apiKey, version) {
  const url = `wss://spark-api.xf-yun.com/${version}/chat`;
  const urlObject = new URL(url);
  const host = urlObject.host;
  const pathname = urlObject.pathname;
  const date = new Date().toGMTString();
  const algorithm = "hmac-sha256";
  const headers = "host date request-line";
  const signatureOrigin = `host: ${host}\ndate: ${date}\nGET ${pathname} HTTP/1.1`;
  const signatureSha = CryptoJS.HmacSHA256(signatureOrigin, apiSecret);
  const signature = CryptoJS.enc.Base64.stringify(signatureSha);
  const authorizationOrigin = `api_key="${apiKey}", algorithm="${algorithm}", headers="${headers}", signature="${signature}"`;
  const authorization = btoa(authorizationOrigin);
  return `${url}?authorization=${authorization}&date=${date}&host=${host}`;
}

/**
 * 验证请求参数
 *
 * @param apiKey apiKey
 * @param prompt 提示信息
 * @param value 单元格的值
 * @returns 返回验证结果
 */
function validateUserPrompt(prompt, value) {
  const responseLiteral = {
    error: false,
    errorMsg: "No error",
    messages: [],
  };

  if (window.systemPrompt && window.systemPrompt.trim() !== "") {
    responseLiteral.messages.push({
      role: "system",
      content: window.systemPrompt,
    });
  }

  if (value === null || value === undefined) value = "";

  let userPrompt = prompt + " " + value;
  if (window.userPrompt && window.userPrompt.trim() !== "") {
    userPrompt = window.userPrompt.replace("{value}", value).replace("{prompt}", prompt);
  }

  if (userPrompt.trim() === "") {
    responseLiteral.error = true;
    responseLiteral.errorMsg = "用户输入为空";
  }
  // 添加用户消息
  responseLiteral.messages.push({
    role: "user",
    content: userPrompt,
  });
  return responseLiteral;
}

/**
 * 填充偏移单元格
 *
 * @param fillOffset 偏移量
 * @param result 填充结果
 * @param invocation 调用参数
 */
async function fillOffsetCell(fillOffset, result, invocation) {
  await Excel.run(async (context) => {
    console.log(invocation.parameterAddresses[1]);
    console.log(invocation.address);
    const [sheetId, cellId] = invocation.address.split("!");
    const invocationCell = context.workbook.worksheets.getItem(sheetId).getRange(cellId);
    const fillOffsetCell = invocationCell.getOffsetRange(0, fillOffset);
    fillOffsetCell.values = [[result]];
    fillOffsetCell.format.autofitColumns();
    await context.sync();
  });
}
