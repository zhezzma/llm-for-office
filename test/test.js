var CryptoJS = require("crypto-js");
const WebSocket = require("ws");
async function gpt(prompt, target) {
  const url = "https://openai-forward-s4pz.onrender.com/v1/chat/completions";
  const apiKey = "";
  if (!prompt) return "prompt 不能为空";
  if (!apiKey) return "apiKey 未设置";
  if (target === null || target === undefined) target = "";
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
            content: prompt + " " + target,
          },
        ],
      }),
    });
    const json = await response.json();
    if (json.error) {
      throw new Error(json.error.message);
    }
    console.log(json.choices[0].message.content);
    return json.choices[0].message.content;
  } catch (error) {
    console.log(error.message);
    return error.message;
  }
}

/**
 * 使用星火生成你想要的数据
 * @customfunction SPARK
 * @param {string} prompt 咒语
 * @param {string} [target] 单元格,如果省略,则为空
 * @returns {string} Result
 */
async function spark(prompt, target) {
  let version = "v3.5";
  let domain = "generalv3.5";
  let parts = "".split(".");
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
              content: prompt + " " + target,
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
      total_res += jsonData.payload.choices.text[0].content;
      // 提问失败
      if (jsonData.header.code !== 0) {
        total_res = jsonData.header.message;
        break;
      }
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
  return total_res;
}

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


async function glm(prompt, target) {
  const url = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
  const apiKey = generateGLMToken("",3600);
  if (!prompt) return "prompt 不能为空";
  if (!apiKey) return "apiKey 未设置";
  if (target === null || target === undefined) target = "";
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
            content: prompt + " " + target,
          },
        ],
      }),
    });
    const json = await response.json();
    if (json.error) {
      throw new Error(json.error.message);
    }
    console.log(json.choices[0].message.content);
    return json.choices[0].message.content;
  } catch (error) {
    console.log(error.message);
    return error.message;
  }
}

function generateGLMToken(apikey, expSeconds) {
  const [id, secret] = apikey.split('.');
  const payload = {
      api_key: id,
      exp: Math.floor(Date.now() / 1000) + expSeconds,
      timestamp: Math.floor(Date.now() / 1000),
  };
  const header = {
      alg: 'HS256',
      sign_type: 'SIGN'
  };
  const stringifiedHeader = CryptoJS.enc.Utf8.parse(JSON.stringify(header));
  const stringifiedPayload = CryptoJS.enc.Utf8.parse(JSON.stringify(payload));
  const encodedHeader = CryptoJS.enc.Base64.stringify(stringifiedHeader);
  const encodedPayload = CryptoJS.enc.Base64.stringify(stringifiedPayload);
  const signature = CryptoJS.HmacSHA256(encodedHeader + '.' + encodedPayload, secret);
  const encodedSignature = CryptoJS.enc.Base64.stringify(signature);
  return encodedHeader + '.' + encodedPayload + '.' + encodedSignature;
}

glm("你好，我是小刘。")
//spark("你好，我是小刘。");
// spark("你好，我是小刘。").then(result => {
//   console.log(result);
//   // Do something with 'a'
// }).catch(error => {
//   // Handle any errors
// });
//gpt("你好，我是小刘。")
