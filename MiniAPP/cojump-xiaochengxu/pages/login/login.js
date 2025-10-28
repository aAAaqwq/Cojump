// pages/login/login.js - COJUMP康复训练系统登录页（微信授权登录）
const app = getApp()

Page({
  data: {
    userInfo: null, // 用户信息
  },

  /**
   * 生命周期函数--监听页面加载
   */
  // onLoad(options) {
  //   // 检查是否已经登录
  //   this.checkLoginStatus()
  // },

  onShow() {
    this.checkLoginStatus()
  },

  /**
   * 检查登录状态
   */
  async checkLoginStatus() {
    // 获取用户信息
    const userInfo = await this.getUserInfo()
    //如果用户信息存在，则跳转到首页
    if (userInfo) {
      this.setData({
        userInfo: userInfo
      })
      
      // 确保缓存已设置
      wx.setStorageSync('userInfo', userInfo)
      
      wx.reLaunch({
        url: '/pages/index/index'
      })
    }
  },

  /**
   * 获取用户信息，先查缓存，缓存没有则查询数据库中是否存在用户信息
   */
  async getUserInfo() {
    return new Promise((resolve, reject) => {
      // 1. 从本地缓存中获取用户信息
      const cachedUserInfo = wx.getStorageSync('userInfo')
      if (cachedUserInfo) {
        console.log('从本地缓存中获取到用户信息:', cachedUserInfo)
        resolve(cachedUserInfo)
        return
      }
      
      // 2. 如果本地缓存中没有用户信息，则查询数据库中是否存在用户信息
      wx.cloud.callFunction({
        name: 'getUserInfo',
        data: {
          openId: app.globalData.openId
        }
      }).then(res => {
        if (res.result && res.result.success) {
          const userInfo = res.result.userInfo
          
          // 保存到缓存
          wx.setStorageSync('userInfo', userInfo)
          console.log('从数据库中获取到用户信息并已缓存:', userInfo)
          
          resolve(userInfo)
        } else {
          console.log('数据库中不存在用户信息')
          resolve(null)
        }
      }).catch(error => {
        console.error('获取用户信息失败', error)
        resolve(null)
      })
    })
  },

  /**
   * 微信一键登录
   * 直接登录，不获取用户授权信息
   */
  wechatLogin() {
    // 直接登录，使用默认用户信息
    wx.showLoading({
      title: '登录中...',
      mask: true
    })
  
    // 调用云函数进行登录
    wx.cloud.callFunction({
      name: 'wechatLogin',
      data: {
        openId: app.globalData.openId,
        userInfo: this.data.userInfo // 首次登录不获取用户信息，使用默认值
      }
    }).then(res => {
      wx.hideLoading()
      console.log('云函数返回:', res)
      
      if (res.result.success) {
        // 缓存用户信息到本地
        const savedUserInfo = res.result.userInfo
        wx.setStorageSync('userInfo', savedUserInfo)
        
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
      console.error('登录云函数调用失败', err)
      
      wx.showToast({
        title: '登录失败，请重试',
        icon: 'none',
        duration: 2000
      })
    })
  },

  /**
   * 使用用户信息进行登录
   */
  loginWithUserInfo(userInfo) {
    wx.showLoading({
      title: '登录中...',
      mask: true
    })
  
    // 调用云函数进行登录
    wx.cloud.callFunction({
      name: 'wechatLogin',
      data: {
        openId: app.globalData.openId,
        userInfo: this.userInfo
      }
    }).then(res => {
      wx.hideLoading()
      console.log('云函数返回:', res)
      
      if (res.result.success) {
        // 保存用户信息到本地
        const savedUserInfo = res.result.userInfo
        wx.setStorageSync('userInfo', savedUserInfo)
        
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
      console.error('登录云函数调用失败', err)
      
      // 开发阶段提示 - 云函数未配置时的临时处理
      wx.showModal({
        title: '开发提示',
        content: '云函数未配置或调用失败\n是否使用本地模拟登录？',
        confirmText: '本地登录',
        cancelText: '取消',
        success: (modalRes) => {
          if (modalRes.confirm) {
            // 本地模拟登录（开发阶段）
            const mockUserInfo = {
              _id: 'mock_' + Date.now(),
              openId: app.globalData.openId,
              nickName: userInfo.nickName || '微信用户',
              avatarUrl: userInfo.avatarUrl || '',
              createTime: new Date().getTime(),
              lastLoginTime: new Date().getTime()
            }
            wx.setStorageSync('userInfo', mockUserInfo)
            
            wx.showToast({
              title: '本地登录成功',
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
