# COJUMP 小程序云函数配置指南

## 📋 数据库设计

### users 集合（简化版）

```javascript
{
  _id: "自动生成的用户ID",
  openId: "微信用户的openId（唯一标识）",
  nickName: "用户昵称（默认：微信昵称）",
  avatarUrl: "头像URL（云存储地址：cloud://xxx/avatars/xxx.jpg）",
  createTime: Date,  // 注册时间（服务端时间）
  lastLoginTime: Date  // 最后登录时间（服务端时间）
}
```

#### 索引设置
- **主键**：_id（自动）
- **唯一索引**：openId（建议创建，避免重复）

---

## ☁️ 云函数列表

### 1. getOpenID（已有）
**功能**：获取用户的openId

**调用位置**：`app.js` 第18行

**状态**：✅ 已实现

---

### 2. wechatLogin（新增）
**功能**：微信一键登录

**路径**：`cloudfunctions/wechatLogin/`

**文件结构**：
```
wechatLogin/
├── index.js      （已创建）
├── config.json   （已创建）
└── package.json  （已创建）
```

**核心逻辑**：
```javascript
1. 接收微信用户信息和openId
2. 查询users集合，检查用户是否存在
3. 如果是新用户：
   - 创建新记录（openId, nickName, avatarUrl, createTime, lastLoginTime）
   - 返回用户信息和token
4. 如果是老用户：
   - 更新lastLoginTime和基本信息
   - 返回更新后的用户信息和token
```

**部署命令**：
```bash
# 在微信开发者工具中
右键 cloudfunctions/wechatLogin → 上传并部署：云端安装依赖
```

---

### 3. updateUserInfo（新增）
**功能**：更新用户信息（昵称、头像）

**路径**：`cloudfunctions/updateUserInfo/`

**文件结构**：
```
updateUserInfo/
├── index.js      （已创建）
├── config.json   （已创建）
└── package.json  （已创建）
```

**核心逻辑**：
```javascript
1. 接收openId、nickName、avatarUrl
2. 查找用户记录
3. 如果有新头像：
   - 删除旧头像（云存储）
   - 更新avatarUrl字段
4. 更新nickName（如果提供）
5. 返回更新后的用户信息
```

**部署命令**：
```bash
# 在微信开发者工具中
右键 cloudfunctions/updateUserInfo → 上传并部署：云端安装依赖
```

---

## 🖼️ 云存储配置

### 存储路径设计
```
cloud://你的环境ID/
└── avatars/
    ├── {openId}_1698234567890.jpg
    ├── {openId}_1698234568901.jpg
    └── ...
```

### 命名规则
- **格式**：`{openId}_{timestamp}.jpg`
- **优点**：
  - openId确保唯一性
  - timestamp防止缓存问题
  - 便于追溯和管理

### 存储权限设置
在云开发控制台设置：
```
路径：avatars/*
读权限：所有用户
写权限：仅创建者
```

---

## 🚀 部署步骤

### 第一步：创建数据库集合
1. 打开微信开发者工具
2. 进入"云开发控制台"
3. 点击"数据库"
4. 点击"添加集合"
5. 集合名称：`users`
6. 权限设置：仅创建者可读写

### 第二步：创建索引（可选但推荐）
```javascript
// 在users集合中创建索引
字段：openId
类型：唯一索引（unique）
```

### 第三步：部署云函数

#### 部署 wechatLogin
```bash
1. 右键 cloudfunctions/wechatLogin
2. 点击"上传并部署：云端安装依赖"
3. 等待部署完成（约30秒）
4. 查看日志确认部署成功
```

#### 部署 updateUserInfo
```bash
1. 右键 cloudfunctions/updateUserInfo
2. 点击"上传并部署：云端安装依赖"
3. 等待部署完成（约30秒）
4. 查看日志确认部署成功
```

### 第四步：配置云存储
```bash
1. 进入"云开发控制台"
2. 点击"存储"
3. 如果是首次使用，会自动创建存储空间
4. 设置存储权限（可选）
```

---

## 🔄 完整流程演示

### 登录流程
```
1. 用户打开小程序 → 进入登录页
   ↓
2. 点击"微信授权登录"按钮
   ↓
3. 微信弹出授权弹窗（系统级）
   ├─ 显示：头像
   ├─ 显示：昵称
   └─ 按钮：允许 / 拒绝
   ↓
4. 用户点击"允许"
   ↓
5. 获取用户信息（nickName, avatarUrl）
   ↓
6. 调用云函数 wechatLogin
   ├─ 传入：userInfo, openId
   ├─ 云函数处理（查询/创建用户）
   └─ 返回：success, userInfo, token
   ↓
7. 保存到本地存储
   ├─ wx.setStorageSync('userInfo', userInfo)
   └─ wx.setStorageSync('token', token)
   ↓
8. 跳转到首页
```

### 修改信息流程
```
1. 首页点击头像 → 进入个人信息页
   ↓
2. 点击"修改个人信息"按钮
   ↓
3. 底部弹出编辑弹窗
   ↓
4. 选择头像
   ├─ 点击头像区域
   ├─ 微信弹出选择框（系统级）
   ├─ 用户选择/拍摄照片
   ├─ 本地预览显示
   └─ 标记为待上传
   ↓
5. 编辑昵称
   ├─ 点击输入框
   ├─ 输入新昵称
   └─ 显示字数（x/12）
   ↓
6. 点击"确认修改"
   ↓
7. 【如果有新头像】
   ├─ 压缩图片（quality: 80）
   ├─ 上传到云存储（显示进度）
   ├─ 获取云存储URL（fileID）
   └─ 继续下一步
   ↓
8. 调用云函数 updateUserInfo
   ├─ 传入：openId, nickName, avatarUrl
   ├─ 云函数更新数据库
   └─ 返回：success, userInfo
   ↓
9. 更新本地存储
   ↓
10. 关闭弹窗，刷新页面
    ↓
11. 显示成功提示
```

---

## 🎨 UI效果说明

### 登录页授权弹窗
- **类型**：微信系统级弹窗（无法自定义）
- **内容**：
  ```
  ┌──────────────────────┐
  │   [头像预览]          │
  │   获取你的昵称、头像   │
  │                      │
  │   [允许]  [拒绝]     │
  └──────────────────────┘
  ```

### 个人信息修改弹窗（自定义）
- **类型**：底部半屏弹窗
- **特点**：
  - 渐变紫色标题栏
  - 大头像预览（160rpx）
  - 带图标的输入框
  - 实时字数统计
  - 上传进度显示
  - 确认/取消双按钮

---

## 💾 云存储最佳实践

### 图片压缩
```javascript
wx.compressImage({
  src: tempFilePath,
  quality: 80,  // 压缩质量 0-100
  success: (res) => {
    // 使用压缩后的图片
    uploadToCloud(res.tempFilePath)
  }
})
```

### 上传进度监听
```javascript
const uploadTask = wx.cloud.uploadFile({...})

uploadTask.onProgressUpdate((res) => {
  console.log('上传进度', res.progress)
  console.log('已上传', res.totalBytesSent)
  console.log('总大小', res.totalBytesExpectedToSend)
})
```

### 删除旧文件
```javascript
// 删除云存储文件
wx.cloud.deleteFile({
  fileList: ['cloud://xxx/avatars/old_avatar.jpg']
}).then(res => {
  console.log('删除成功', res.fileList)
}).catch(err => {
  console.error('删除失败', err)
})
```

---

## 🔐 安全性建议

### 1. 文件大小限制
```javascript
// 在上传前检查文件大小
wx.getFileSystemManager().getFileInfo({
  filePath: tempFilePath,
  success: (res) => {
    const fileSize = res.size
    if (fileSize > 2 * 1024 * 1024) {  // 2MB
      wx.showToast({
        title: '图片过大，请选择小于2MB的图片',
        icon: 'none'
      })
      return
    }
    // 继续上传
  }
})
```

### 2. 文件类型验证
```javascript
// 只允许图片
wx.chooseMedia({
  count: 1,
  mediaType: ['image'],  // 只允许图片
  sourceType: ['album', 'camera']
})
```

### 3. 昵称验证
```javascript
// 云函数中验证昵称
if (nickName.length > 12) {
  return { success: false, message: '昵称不能超过12个字符' }
}

// 可选：敏感词过滤
const forbiddenWords = ['违规词1', '违规词2']
if (forbiddenWords.some(word => nickName.includes(word))) {
  return { success: false, message: '昵称包含敏感词' }
}
```

---

## 🧪 测试步骤

### 测试登录流程
1. ✅ 清除缓存（Storage）
2. ✅ 重新打开小程序
3. ✅ 应该显示登录页
4. ✅ 点击"微信授权登录"
5. ✅ 允许授权
6. ✅ 检查是否跳转到首页
7. ✅ 检查首页右上角是否显示头像

### 测试修改信息
1. ✅ 点击首页右上角头像
2. ✅ 进入个人信息页
3. ✅ 点击"修改个人信息"
4. ✅ 弹窗从底部滑入
5. ✅ 点击头像，选择新头像
6. ✅ 修改昵称
7. ✅ 点击"确认修改"
8. ✅ 观察上传进度
9. ✅ 检查是否更新成功

---

## 📱 云开发环境配置

### 获取环境ID
1. 打开"云开发控制台"
2. 顶部显示环境ID（如：`cloudbase-xxx`）
3. 复制环境ID

### 更新 app.js
```javascript
wx.cloud.init({
  env: "你的云环境ID"  // 替换为实际的环境ID
})
```

---

## 🎯 关键代码片段

### 上传头像到云存储
```javascript
const cloudPath = `avatars/${openId}_${Date.now()}.jpg`

const uploadTask = wx.cloud.uploadFile({
  cloudPath: cloudPath,
  filePath: tempFilePath,
  success: (res) => {
    console.log('上传成功', res.fileID)
    // res.fileID 就是云存储URL
  }
})

// 监听进度
uploadTask.onProgressUpdate((res) => {
  console.log('进度:', res.progress + '%')
})
```

### 调用云函数
```javascript
wx.cloud.callFunction({
  name: 'updateUserInfo',
  data: {
    openId: app.globalData.openId,
    nickName: '新昵称',
    avatarUrl: 'cloud://xxx/avatars/xxx.jpg'
  }
}).then(res => {
  if (res.result.success) {
    // 更新成功
    wx.setStorageSync('userInfo', res.result.userInfo)
  }
})
```

---

## 🐛 常见问题

### Q1: 云函数调用失败？
**A**: 检查以下几点：
- 云环境ID是否正确
- 云函数是否已部署
- 网络是否正常
- 查看云函数日志

### Q2: 头像上传失败？
**A**: 可能原因：
- 云存储未开通
- 文件过大（>10MB）
- 网络不稳定
- 云存储权限设置问题

### Q3: 旧头像删除失败？
**A**: 
- 检查文件是否存在
- 检查权限设置
- 不影响主流程，可忽略

### Q4: 本地预览头像显示失败？
**A**:
- 临时文件路径可能失效
- 使用 `wx.getFileSystemManager()` 读取
- 确保在选择后立即预览

---

## 📊 数据库查询示例

### 查询单个用户
```javascript
db.collection('users')
  .where({ openId: 'xxx' })
  .get()
```

### 更新用户信息
```javascript
db.collection('users')
  .doc(userId)
  .update({
    data: {
      nickName: '新昵称',
      avatarUrl: 'cloud://xxx',
      lastLoginTime: db.serverDate()
    }
  })
```

---

## ✅ 验收标准

### 登录功能
- [x] 首次登录自动创建用户记录
- [x] 保存用户信息到本地存储
- [x] 显示微信头像和昵称
- [x] 登录后跳转到首页

### 个人信息修改
- [x] 弹窗美观，从底部滑入
- [x] 可以选择新头像
- [x] 头像本地即时预览
- [x] 可以修改昵称
- [x] 显示字数统计（x/12）
- [x] 显示上传进度
- [x] 头像存储到云存储
- [x] 自动删除旧头像
- [x] 更新成功后刷新页面

### 首页显示
- [x] 右上角显示用户头像
- [x] 点击头像进入个人信息页

---

## 🎨 UI特点总结

1. **登录页**：
   - COJUMP品牌突出
   - 微信绿色登录按钮
   - 简洁优雅

2. **个人信息页**：
   - 渐变顶部背景
   - 大头像展示
   - 信息卡片化
   - 圆形返回按钮

3. **编辑弹窗**：
   - 底部半屏设计
   - 渐变标题栏
   - 头像预览+角标
   - 字数实时统计
   - 上传进度条
   - 双按钮布局

---

**所有功能已完成开发，请部署云函数后测试！** 🎉

