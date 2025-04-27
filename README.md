# LLM for Office

https://llmoffice.godgodgame.com/taskpane.html

一个将大型语言模型(LLM)集成到Microsoft Office中的插件，为用户提供智能写作、数据分析和自然语言查询等功能。

## 功能特点

- 支持多种LLM模型集成 (GPT、星火、GLM等)
- 提供Excel自定义函数用于文本生成
- 可配置的系统提示和用户提示模板
- 支持API并发控制
- 支持结果过滤处理

## 开发环境要求

- Node.js (>=16 <21)
- npm (>=7 <11)
- Microsoft Excel

## 快速开始

1. 克隆项目并安装依赖:

```bash
git clone https://github.com/zhezzma/llm-for-office.git
cd llm-for-office
npm install
```

2. 开发模式运行:

```bash
npm run dev-server
```

3. 生产环境构建:

```bash
npm run build
```

## Excel函数使用示例

```excel
=G.GPT("为道具生成描述", A1)
```

## 配置说明

在插件任务窗格中可以配置以下参数：

- 系统提示 (System Prompt)
- 用户提示格式 (User Prompt Format)
- 过滤符号 (Filter Pattern)
- API并发数 (Semaphore Count)
- GPT API地址
- GPT API密钥
- GPT模型选择

## 环境变量配置

项目支持使用 `.env` 文件进行环境变量配置。在项目根目录创建 `.env` 文件，可以配置以下参数：

```
# API Configuration
PRODUCTION_URL=https://your-production-url.com/

```

## 部署

项目使用webpack进行构建，生产环境部署地址可以通过 `.env` 文件中的 `PRODUCTION_URL` 变量进行配置：

```javascript
// webpack.config.js
const urlProd = process.env.PRODUCTION_URL || "https://llmoffice.godgodgame.com/";
```

## 开发命令

- `npm run build` - 生产环境构建
- `npm run build:dev` - 开发环境构建
- `npm run dev-server` - 启动开发服务器
- `npm run start` - 启动插件
- `npm run start:desktop` - 在桌面Excel中启动插件
- `npm run start:web` - 在Web Excel中启动插件
- `npm run lint` - 代码检查
- `npm run lint:fix` - 自动修复代码问题

## 许可证

MIT

## 相关链接

- [官网](https://www.godgodgame.com)
- [GitHub](https://github.com/zhezzma)