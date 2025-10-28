// pages/login/login.js - COJUMP康复训练系统登录页（微信授权登录）
const app = getApp()

Page({
  data: {
    userInfo: null, // 用户信息
  },

  /**
   * 生命周期函数--监听页面加载
   */
  onLoad(options) {
    // 检查是否已经登录
    this.checkLoginStatus()
  },

  /**
   * 检查登录状态
   */
  checkLoginStatus() {
    const userInfo = wx.getStorageSync('userInfo')
    if (userInfo) {
      this.setData({ userInfo })
      // 如果已登录，跳转到首页
      wx.reLaunch({
        url: '/pages/index/index'
      })
    }
  },

  /**
   * 微信授权登录
   */
  wechatLogin(e) {
    console.log('微信登录授权', e)
    
    // 获取用户信息
    if (e.detail.userInfo) {
      wx.showLoading({
        title: '登录中...',
        mask: true
      })

      // 调用云函数进行登录
      wx.cloud.callFunction({
        name: 'wechatLogin', // 云函数名称，需要您创建
        data: {
          userInfo: e.detail.userInfo,
          openId: app.globalData.openId
        }
      }).then(res => {
        wx.hideLoading()
        
        if (res.result.success) {
          // 保存用户信息
          const userInfo = res.result.userInfo
          wx.setStorageSync('userInfo', userInfo)
          wx.setStorageSync('token', res.result.token)
          
          wx.showToast({
            title: '登录成功',
            icon: 'success'
          })

          // 跳转到首页
          setTimeout(() => {
            wx.reLaunch({
              url: '/pages/index/index'
            })
          }, 1500)
        } else {
          wx.showToast({
            title: res.result.message || '登录失败',
            icon: 'none'
          })
        }
      }).catch(err => {
        wx.hideLoading()
        console.error('微信登录失败:', err)
        
        // 开发阶段提示 - 云函数未配置时的临时处理
        wx.showModal({
          title: '提示',
          content: '登录功能需要配置云函数\n开发阶段演示：直接跳转首页',
          confirmText: '继续',
          cancelText: '取消',
          success: (modalRes) => {
            if (modalRes.confirm) {
              // 保存用户信息（开发阶段临时方案）
              const userInfo = {
                ...e.detail.userInfo,
                openId: app.globalData.openId,
                createTime: new Date().getTime()
              }
              wx.setStorageSync('userInfo', userInfo)
              
              wx.showToast({
                title: '登录成功',
                icon: 'success'
              })
              
              setTimeout(() => {
                wx.reLaunch({
                  url: '/pages/index/index'
                })
              }, 1500)
            }
          }
        })
      })
    } else {
      // 用户拒绝授权
      wx.showToast({
        title: '需要授权才能使用',
        icon: 'none',
        duration: 2000
      })
    }
  },

  /**
   * 显示用户协议
   */
  showAgreement() {
    wx.showModal({
      title: '用户协议',
      content: 'COJUMP智能康复训练系统用户协议\n\n1. 用户在使用本系统前应仔细阅读本协议\n2. 用户使用本系统即表示同意本协议的全部内容\n3. 本系统仅用于康复训练辅助，不替代专业医疗建议\n4. 用户应妥善保管账号信息\n5. 本系统会保护用户隐私和数据安全\n\n如有疑问，请联系客服。',
      showCancel: false,
      confirmText: '我知道了'
    })
  },

  /**
   * 显示隐私政策
   */
  showPrivacy() {
    wx.showModal({
      title: '隐私政策',
      content: 'COJUMP智能康复训练系统隐私政策\n\n1. 我们重视用户隐私保护\n2. 仅收集必要的用户信息用于提供服务\n3. 不会向第三方泄露用户信息\n4. 用户数据采用加密存储和传输\n5. 用户有权查看、修改或删除个人信息\n6. 我们会及时告知隐私政策的重大变更\n\n如有疑问，请联系客服。',
      showCancel: false,
      confirmText: '我知道了'
    })
  },

  /**
   * 分享配置
   */
  onShareAppMessage() {
    return {
      title: 'COJUMP智能康复训练系统',
      path: '/pages/login/login',
      imageUrl: '/image/KANG_YU.png'
    }
  }
})
