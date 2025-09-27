// app.js
App({
  globalData: {
    // 调试打印信息
    debugFlag: true,
    // 开发者服务器基地址
    developerServerBaseUrl: 'https://localhost',
    // 保存用户openid
    openid: '',
    // 设备id
    dev_id: 'cojump_device1',
    // 保存用户登录信息
    userInfo: null,
  },
  onLaunch() {
    if (this.globalData.debugFlag) {
      console.log('app.js===>onLaunch() do start!')
    }
    //初始化云函数
    wx.cloud.init({
      //云环境id
      env: "cloudbase-7gky5kj9737551d9"
    })
    // 展示本地存储能力
    const logs = wx.getStorageSync('logs') || []
    logs.unshift(Date.now())
    wx.setStorageSync('logs', logs)

    if (this.globalData.debugFlag) {
      console.log('app.js===>onLaunch() do end!')
    }
  },
  OnShow() {

  },
  OnHide() {
    // To do
  },
  onError() {
    //To do
  },
})
