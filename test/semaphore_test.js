
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




  const semaphore = new Semaphore(2);
  async function AsyncFunction(index) {
    await semaphore.acquire();
    try {
      // 在这里执行您的异步任务
      // ... 执行的函数 ..
      console.log("执行函数"+index);
      // 模拟异步任务需要一秒钟
      await new Promise(resolve => setTimeout(resolve, 1000));
      return "xxx"+index; // 返回您想要的任何内容
    } finally {
        console.log("释放"+index)
      semaphore.release();
    }
  }
  // 示范用法：
  async function test() {
    const results = await Promise.all([
      AsyncFunction(1),
      AsyncFunction(2),
      AsyncFunction(3),
      AsyncFunction(4),
    ]);
    console.log(results); // 打印出 ['xxx', 'xxx', 'xxx', 'xxx']
  }
  test();