/* global clearInterval, console, setInterval */


window.glmKey = localStorage.getItem('glm-key');
window.gptKey = localStorage.getItem('gpt-key');
/**
 * 使用chatglm生成你想要的数据
 * @customfunction GLM
 * @param {string} prompt 咒语
 * @param {string} [target] 单元格,如果省略,则为空
 * @returns {string} Result  
 */
export async function glm(prompt, target) {
  const url = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
  const apiKey = window.glmKey;
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


/**
 * 使用chatgpt生成你想要的数据
 * @customfunction GPT
 * @param {string} prompt 咒语
 * @param {string} [target] 单元格,如果省略,则为空
 * @returns {string} Result 
 */
export async function gpt(prompt, target) {
  const url = "https://openai-forward-s4pz.onrender.com/v1/chat/completions";
  const apiKey = window.gptKey;
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