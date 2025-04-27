/* global clearnumbererval, console, setnumbererval */
//https://learn.microsoft.com/zh-cn/office/dev/add-ins/excel/

class Semaphore {
  constructor(count) {
    this.count = count;
    this.waiting = [];
  }
  acquire() {
    if (this.count > 0) {
      this.count--;
      return Promise.resolve(true);
    }
    return new Promise((resolve) => {
      this.waiting.push(resolve);
    });
  }
  release() {
    if (this.waiting.length > 0) {
      const resolve = this.waiting.shift();
      resolve(true);
    } else {
      this.count++;
    }
  }
}

let semaphore = new Semaphore(100);

Office.onReady(() => {
  if (window.semaphoreCount) {
    semaphore.count = parseInt(window.semaphoreCount);
  }
});

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
  await semaphore.acquire();
  let result = "";

  try {
    const model = window.gptModel || "gpt-4";
    const url = window.gptUrl || "https://api.openai.com/v1/chat/completions";
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
    // Remove content between <think> and </think> tags (including the tags)
    result = result.replace(/<think>[\s\S]*?<\/think>/g, '');
  } catch (error) {
    result = error.message;
  } finally {
    semaphore.release();
  }

  console.log(result);
  if (fillOffset != 0) {
    await fillOffsetCell(fillOffset, result, invocation);
  }
  return result;
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
    let filteredText = result;
    let filterPatternInput = window.filterPattern;
    // 如果过滤模式不为空，则执行正则替换
    if (filterPatternInput) {
      // 将输入的字符以|为分隔符进行分割
      const patterns = filterPatternInput.split("|");
      // 创建一个用于匹配所有输入符号的正则表达式
      const regex = new RegExp(
        patterns
          .map(function (character) {
            // 对特殊的正则表达式字符进行转义
            return character.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          })
          .join("|"),
        "g"
      );
      // 使用正则表达式替换掉所有匹配的字符
      filteredText = filteredText.replace(regex, "");
    }
    fillOffsetCell.values = [[filteredText]];
    fillOffsetCell.format.autofitColumns();
    await context.sync();
  });
}
