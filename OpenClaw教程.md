# OpenClaw 入门教程

> 本教程适用于 2026 年 3 月最新版本

---

## 什么是 OpenClaw？

OpenClaw 是一个**开源个人 AI 助手**，可以：
- 在多种平台运行（macOS/Linux/Windows）
- 通过各种聊天软件对话（QQ、Telegram、Discord、WhatsApp 等）
- 拥有持久记忆，记住你的偏好
- 自主执行任务（发邮件、管理日历、自动化操作等）

**官网**：https://openclaw.ai
**文档**：https://docs.openclaw.ai

---

## 安装

### 环境要求
- Node.js 22+
- macOS / Linux / Windows（推荐 WSL2）

### 一键安装

**macOS / Linux / WSL2：**
```bash
curl -fsSL https://openclaw.ai/install.sh | bash
```

**Windows (PowerShell)：**
```powershell
iwr -useb https://openclaw.ai/install.ps1 | iex
```

---

## 初始化配置

### 运行引导向导
```bash
openclaw onboard --install-daemon
```

向导会帮你配置：
1. API 认证（支持 Anthropic、OpenAI、MiniMax 等）
2. 网关设置
3. 聊天通道（QQ/Telegram/Discord 等）
4. 技能安装

### 检查状态
```bash
openclaw gateway status
```

### 打开控制台
```bash
openclaw dashboard
```

然后访问：`http://127.0.0.1:18789/`

---

## 配置聊天通道

### QQ 配置
```bash
openclaw channels add --channel qqbot --token "AppID:ClientSecret"
```

例如：
```bash
openclaw channels add --channel qqbot --token "1903480275:你的ClientSecret"
```

### Telegram 配置
```bash
openclaw channels add --channel telegram --token "你的BOT_TOKEN"
```

### Discord 配置
```bash
openclaw channels add --channel discord --token "你的BOT_TOKEN"
```

---

## 安装技能

### 查看可用技能
```bash
openclaw skills list
```

### 安装技能
```bash
openclaw skills install <skill-name>
```

### 常用技能
- `weather` - 天气查询
- `notion` - Notion 集成
- `obsidian` - Obsidian 笔记
- `github` - GitHub 操作

### 社区技能（ClawHub）
```bash
openclaw plugins install <plugin-name>
```

---

## 常用命令

| 命令 | 说明 |
|------|------|
| `openclaw status` | 查看状态 |
| `openclaw gateway start` | 启动网关 |
| `openclaw gateway stop` | 停止网关 |
| `openclaw dashboard` | 打开控制台 |
| `openclaw configure` | 重新配置 |
| `openclaw channels list` | 查看通道 |

---

## 工作目录

OpenClaw 的工作目录：
- **配置**：`~/.openclaw/` 或 `/opt/openclaw/data/.openclaw/`
- **工作文件**：`~/.openclaw/workspace/`

重要文件：
- `openclaw.json` - 主配置
- `SOUL.md` - AI 人设
- `USER.md` - 用户信息
- `MEMORY.md` - 长期记忆
- `memory/` - 每日记忆

---

## 进阶功能

### 定时任务 (Cron)
```bash
openclaw cron add "0 7 * * *" "天气查询"
```

### 心跳检查 (Heartbeat)
编辑 `HEARTBEAT.md` 文件，设定周期性检查任务

### 记忆系统
- `MEMORY.md` - 长期记忆（重要信息）
- `memory/YYYY-MM-DD.md` - 每日记录

---

## 常见问题

### Q: 如何更新 OpenClaw？
```bash
npm update -g openclaw
```

### Q: 内存占用太高？
OpenClaw 默认会加载多个模型。可以在配置中调整。

### Q: 想用本地模型？
支持 Ollama、LM Studio 等本地模型。

---

## 总结

```
安装 → 初始化 → 配置通道 → 安装技能 → 开始使用
```

一行命令开始你的 AI 助手之旅：
```bash
curl -fsSL https://openclaw.ai/install.sh | bash
```

---

*更新时间：2026年3月*
