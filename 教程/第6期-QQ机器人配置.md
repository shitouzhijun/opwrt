# 第6期：QQ 机器人配置

## 视频信息
- **时长**：15-18 分钟
- **目标**：配置 QQ 机器人

---

## 开场（20秒）

哈喽大家好，欢迎回到 OpenClaw 教程系列。

前5期我们完成了入门基础，今天来讲——**如何配置 QQ 机器人？**

---

## QQ 机器人选项（3分钟）

### 方式一：官方 QQ Bot API（需要企业认证）
- 适合有开发者资质的用户
- 稳定但申请困难

### 方式二：OneBot 协议（推荐）
- go-cqhttp
- Lagrange.Core
- NapCat
- 无需企业认证

---

## 安装步骤（8分钟）

### Step 1：安装 OpenClaw QQ 插件

```bash
openclaw plugins install @tencent-connect/openclaw-qqbot
```

### Step 2：配置通道

```bash
openclaw channels add --channel qqbot --token "AppID:ClientSecret"
```

### Step 3：重启网关

```bash
openclaw gateway restart
```

---

## QQ 开放平台申请（4分钟）

1. 访问 https://open.openservice.cn/
2. 注册开发者
3. 创建应用
4. 获取 AppID 和 ClientSecret
5. 配置权限

---

## 常见问题（2分钟）

### 无法登录？
- 检查 token 是否正确
- 查看日志：`openclaw logs`

### 消息发不出？
- 检查网络
- 确认通道已启用

---

## 下期预告（20秒）

下一期，我们来讲**Telegram/Discord 配置**。

---

*如果觉得有帮助，请点赞、关注！*
