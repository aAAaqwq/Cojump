// pages/userInfo/userInfo.js
const app = getApp()

Page({
  data: {
    userInfo: null,
    createTimeStr: '',
    lastLoginTimeStr: ''
  },

  onLoad(options) {
    this.loadUserInfo()
  },

  // 加载用户信息
  loadUserInfo() {
    const userInfo = wx.getStorageSync('userInfo')
    if (userInfo) {
      // 格式化时间
      const createTimeStr = this.formatTime(userInfo.createTime)
      const lastLoginTimeStr = this.formatTime(userInfo.lastLoginTime)

      this.setData({
        userInfo,
        createTimeStr,
        lastLoginTimeStr
      })
    } else {
      // 如果没有用户信息，返回登录页
      wx.reLaunch({
        url: '/pages/login/login'
      })
    }
  },

  // 格式化时间
  formatTime(timestamp) {
    if (!timestamp) return '未知'
    
    const date = new Date(timestamp)
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    const hour = String(date.getHours()).padStart(2, '0')
    const minute = String(date.getMinutes()).padStart(2, '0')

    return `${year}-${month}-${day} ${hour}:${minute}`
  },

  // 更新用户信息
  updateUserInfo(e) {
    console.log('更新用户信息', e)
    
    if (e.detail.userInfo) {
      wx.showLoading({
        title: '更新中...',
        mask: true
      })

      // 调用云函数更新用户信息
      wx.cloud.callFunction({
        name: 'updateUserInfo',
        data: {
          userInfo: e.detail.userInfo,
          openId: app.globalData.openId
        }
      }).then(res => {
        wx.hideLoading()
        
        if (res.result.success) {
          // 更新本地存储
          const newUserInfo = res.result.userInfo
          wx.setStorageSync('userInfo', newUserInfo)
          
          // 刷新页面数据
          this.loadUserInfo()
          
          wx.showToast({
            title: '更新成功',
            icon: 'success'
          })
        } else {
          wx.showToast({
            title: res.result.message || '更新失败',
            icon: 'none'
          })
        }
      }).catch(err => {
        wx.hideLoading()
        console.error('更新用户信息失败:', err)
        
        // 开发阶段提示
        wx.showModal({
          title: '提示',
          content: '更新功能需要配置云函数\n开发阶段演示：直接更新本地数据',
          confirmText: '继续',
          success: (modalRes) => {
            if (modalRes.confirm) {
              // 更新本地数据（开发阶段临时方案）
              const updatedUserInfo = {
                ...this.data.userInfo,
                ...e.detail.userInfo,
                lastLoginTime: new Date().getTime()
              }
              wx.setStorageSync('userInfo', updatedUserInfo)
              this.loadUserInfo()
              
              wx.showToast({
                title: '本地更新成功',
                icon: 'success'
              })
            }
          }
        })
      })
    } else {
      wx.showToast({
        title: '需要授权才能更新',
        icon: 'none'
      })
    }
  },

  // 退出登录
  logout() {
    wx.showModal({
      title: '确认退出',
      content: '确定要退出登录吗？',
      confirmColor: '#e74c3c',
      success: (res) => {
        if (res.confirm) {
          // 清除本地存储
          wx.removeStorageSync('userInfo')
          wx.removeStorageSync('token')
          
          wx.showToast({
            title: '已退出登录',
            icon: 'success'
          })
          
          // 返回登录页
          setTimeout(() => {
            wx.reLaunch({
              url: '/pages/login/login'
            })
          }, 1500)
        }
      }
    })
  },

  // 返回
  goBack() {
    wx.navigateBack()
  }
})

