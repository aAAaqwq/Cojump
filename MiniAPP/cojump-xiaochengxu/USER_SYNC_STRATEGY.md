# 用户信息同步策略说明

## 策略概述

根据您的需求，我们实现了以下用户信息同步策略：

1. **优先从本地缓存获取用户信息**
2. **缓存没有查db**
3. **更新后删除本地缓存并重新加载**

## 实现方式

### 1. 用户信息加载流程

#### pages/userInfo/userInfo.js
```javascript
loadUserInfo() {
  // 1. 优先从本地缓存获取
  const cachedUserInfo = wx.getStorageSync('userInfo')
  
  if (cachedUserInfo) {
    // 使用缓存数据
    this.setData({ userInfo: cachedUserInfo })
    return
  }
  
  // 2. 缓存没有，从服务器获取
  this.loadUserInfoFromCloud()
}
```

#### pages/index/index.js
```javascript
loadUserInfo() {
  // 1. 优先从本地缓存获取
  const cachedUserInfo = wx.getStorageSync('userInfo')
  
  if (cachedUserInfo) {
    this.setData({ userInfo: cachedUserInfo })
    return
  }
  
  // 2. 从app全局数据加载
  if (app.globalData.userInfo) {
    this.setData({ userInfo: app.globalData.userInfo })
    wx.setStorageSync('userInfo', app.globalData.userInfo)
    return
  }
  
  // 3. 都不存在，从服务器获取
  this.loadUserInfoFromCloud()
}
```

### 2. 从服务器加载用户信息

#### 新增云函数: getUserInfo
```javascript
// cloudfunctions/getUserInfo/index.js
exports.main = async (event, context) => {
  const { openId } = event
  const wxContext = cloud.getWXContext()
  const actualOpenId = openId || wxContext.OPENID

  // 查询用户信息
  const userQuery = await db.collection('users')
    .where({ openId: actualOpenId })
    .get()

  if (userQuery.data.length === 0) {
    return { success: false, message: '用户不存在' }
  }

  return { success: true, userInfo: userQuery.data[0] }
}
```

### 3. 用户信息更新流程

#### pages/userInfo/userInfo.js
```javascript
// 更新用户信息后
if (res.result && res.result.success) {
  // 清空本地缓存
  wx.removeStorageSync('userInfo')
  
  // 立即从服务器重新获取最新用户信息
  setTimeout(() => {
    this.loadUserInfoFromCloud()
  }, 500)
}
```

## 数据流程图

```
加载用户信息:
┌─────────────────────┐
│ 优先从本地缓存获取 │
└──────────┬──────────┘
           │
           ├─ 有缓存 → 直接使用
           │
           └─ 无缓存 → 调用getUserInfo云函数
                       │
                       ├─ 查询数据库
                       │
                       └─ 保存到本地缓存
                          │
                          └─ 更新页面

更新用户信息:
┌─────────────────────┐
│ 用户编辑个人信息    │
└──────────┬──────────┘
           │
           ├─ 上传头像到云存储
           │
           ├─ 调用updateUserInfo云函数
           │
           ├─ 删除本地缓存
           │
           └─ 重新从服务器获取
              │
              └─ 保存到本地缓存
                 │
                 └─ 更新页面
```

## 优势

1. **性能优化**: 优先使用本地缓存，减少网络请求
2. **数据一致性**: 更新后清除缓存，强制从服务器获取最新数据
3. **用户体验**: 快速加载，无感知的自动同步
4. **容错处理**: 缓存失败时自动降级到服务器获取

## 注意事项

1. **云函数部署**: 需要上传并部署 `getUserInfo` 云函数
2. **缓存失效**: 更新用户信息后会清除缓存，下次加载时会从服务器获取
3. **并发处理**: 多个页面同时加载时会优先使用缓存，避免重复请求

## 如何测试

1. 首次登录后，用户信息会缓存到本地
2. 修改个人信息后，本地缓存会被清除
3. 刷新页面时，会从服务器重新获取最新数据
4. 再次打开应用时，会优先使用本地缓存

