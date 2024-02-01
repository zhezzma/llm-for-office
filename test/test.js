
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


  gpt("你好，我是小刘。")