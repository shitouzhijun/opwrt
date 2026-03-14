# 第7期：Telegram/Discord 配置

## 视频信息
- **时长**：10-15 分钟
- **目标**：配置 Telegram 和 Discord 机器人

---

## 开场（20秒）

哈喽大家好，欢迎回到 OpenClaw 教程系列。

上期我们配置好了 QQ，今天来讲——**如何配置 Telegram 和 Discord？**

---

## Telegram 配置（5分钟）

### Step 1：创建机器人

1. 打开 Telegram
2. 搜索 @BotFather
3. 发送 `/newbot`
4. 设置名称和用户名
5. 获取 **API Token**

### Step 2：配置 OpenClaw

```bash
openclaw channels add --channel telegram --token "你的TOKEN"
```

### Step 3：启动对话

- 搜索你的机器人用户名
- 发送 `/start`

---

## Discord 配置（5分钟）

### Step 1：创建应用

1. 访问 https://discord.com/developers/applications
2. 创建 New Application
3. Bot → Add Bot
4. 获取 **Token**

### Step 2：配置权限

- 勾选 MESSAGE CONTENT INTENT
- 复制 Token

### Step 3：配置 OpenClaw

```bash
openclaw channels add --channel discord --token "你的TOKEN"
```

### Step 4：邀请机器人

生成邀请链接并添加到服务器

---

## 多平台同时使用（2分钟）

可以同时配置多个通道：

```bash
# 查看通道列表
openclaw channels list
```

---

## 下期预告（20秒）

下一期，我们来讲**WhatsApp 配置**。

---

*如果觉得有帮助，请点赞、关注！*
