/* global clearnumbererval, console, setnumbererval */
//https://learn.microsoft.com/zh-cn/office/dev/add-ins/excel/

const CryptoJS = require("crypto-js");

window.glmKey = localStorage.getItem("glm-key");
window.gptKey = localStorage.getItem("gpt-key");
window.sparkKey = localStorage.getItem("spark-key");
/**
 * 使用chatglm生成你想要的数据
 * @customfunction GLM
 * @param {string} prompt 输入提示
 * @param {string} [source] 原始文本
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function glm(prompt, source, fillOffset, invocation) {
  if (!window.glmKey) return "apiKey 未设置";
  if (!prompt) return "prompt 不能为空";
  const url = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
  const apiKey = generateGLMToken(window.glmKey, 3600);
  if (source === null || source === undefined) source = "";
  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;

  let result = "";

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model: "glm-4",
        messages: [
          {
            role: "user",
            content: prompt + " " + source,
          },
        ],
      }),
    });
    const json = await response.json();
    if (json.error) {
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
    let fillOffsetCell = invocationCell.getOffsetRange(0, fillOffset);
    fillOffsetCell.values = [[result]];
    fillOffsetCell.format.autofitColumns();
    await context.sync();
  });
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
 * 使用chatgpt生成你想要的数据
 * @customfunction GPT
 * @param {string} prompt 输入提示
 * @param {string} [source] 原始文本
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function gpt(prompt, source, fillOffset, invocation) {
  if (!prompt) return "prompt 不能为空";
  if (!window.gptKey) return "apiKey 未设置";
  const url = "https://openai-forward-s4pz.onrender.com/v1/chat/completions";
  const apiKey = window.gptKey;
  if (source === null || source === undefined) source = "";
  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;

  let result = "";

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [
          {
            role: "user",
            content: prompt + " " + source,
          },
        ],
      }),
    });
    const json = await response.json();
    if (json.error) {
      throw new Error(json.error.message);
    }
    result = json.choices[0].message.content;
  } catch (error) {
    console.log(error.message);
    result = error.message;
  }

  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
}

/**
 * 使用星火生成你想要的数据
 * @customfunction SPARK
 * @param {string} prompt 输入提示
 * @param {string} [source] 原始文本
 * @param {number} [fillOffset] 填充单元格的偏移量
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 生成的文本
 * @requiresAddress
 * @requiresParameterAddresses
 */
export async function spark(prompt, source, fillOffset, invocation) {
  if (!prompt) return "prompt 不能为空";
  if (!window.sparkKey) return "apiKey 未设置";
  if (source === null || source === undefined) source = "";
  if (fillOffset === null || fillOffset === undefined) fillOffset = 0;
  let version = "v3.5";
  let domain = "generalv3.5";
  let parts = window.sparkKey.split(".");
  let APPID = parts[0];
  let APISecret = parts[1];
  let APIKey = parts[2];
  let url = getSparkUrl(APISecret, APIKey, version);
  let ttsWS = new WebSocket(url);
  let total_res = "";

  try {
    // 等待连接打开
    await new Promise((resolve) => (ttsWS.onopen = resolve));

    // 发送消息
    let params = {
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
          text: [
            {
              role: "user",
              content: prompt + " " + source,
            },
          ],
        },
      },
    };
    ttsWS.send(JSON.stringify(params));

    // 接收消息
    while (true) {
      let e = await new Promise((resolve) => (ttsWS.onmessage = resolve));
      let jsonData = JSON.parse(e.data);
      // 提问失败
      if (jsonData.header.code !== 0) {
        throw new Error(jsonData.header.message);
      }
      total_res += jsonData.payload.choices.text[0].content;
      // 接收完成
      if (jsonData.header.code === 0 && jsonData.header.status === 2) {
        break;
      }
    }
  } catch (error) {
    total_res = error.message;
  } finally {
    // 清理事件监听器
    ttsWS.onmessage = null;
    ttsWS.onclose = null;
    ttsWS.close();
  }
  console.log(total_res);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return total_res;
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
  var url = `wss://spark-api.xf-yun.com/${version}/chat`;
  var host = "spark-api.xf-yun.com";
  var date = new Date().toGMTString();
  var algorithm = "hmac-sha256";
  var headers = "host date request-line";
  var signatureOrigin = `host: ${host}\ndate: ${date}\nGET /${version}/chat HTTP/1.1`;
  var signatureSha = CryptoJS.HmacSHA256(signatureOrigin, apiSecret);
  var signature = CryptoJS.enc.Base64.stringify(signatureSha);
  var authorizationOrigin = `api_key="${apiKey}", algorithm="${algorithm}", headers="${headers}", signature="${signature}"`;
  var authorization = btoa(authorizationOrigin);
  url = `${url}?authorization=${authorization}&date=${date}&host=${host}`;
  return url;
}
