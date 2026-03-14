# 第8期：WhatsApp 配置

## 视频信息
- **时长**：8-10 分钟
- **目标**：配置 WhatsApp 连接

---

## 开场（20秒）

哈喽大家好，欢迎回到 OpenClaw 教程系列。

上期我们配置好了 Telegram 和 Discord，今天来讲——**如何配置 WhatsApp？**

---

## WhatsApp 配置（5分钟）

### 方式一：二维码登录（推荐）

```bash
openclaw channels add --channel whatsapp
```

会生成二维码，用手机 WhatsApp 扫描即可。

### 方式二：Baileys Web

需要安装：

```bash
openclaw plugins install @openclawed/openclaw-whatsapp
```

---

## 配置步骤（3分钟）

### Step 1：添加通道

```bash
openclaw channels add --channel whatsapp
```

### Step 2：扫描二维码

- 打开手机 WhatsApp
- 设置 → 已连接的设备 → 链接设备
- 扫描屏幕上二维码

### Step 3：验证连接

发送消息测试

---

## 常见问题（2分钟）

### 二维码不显示？
- 检查网络
- 重新生成：`openclaw channels add --channel whatsapp --force`

### 连接断开？
- 定期重新扫描
- 使用长期登录方案

---

## 下期预告（20秒）

下一期我们进入**技能进阶**篇章，讲**天气查询**技能。

---

*如果觉得有帮助，请点赞、关注！*
