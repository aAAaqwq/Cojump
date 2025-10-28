// pages/userInfo/userInfo.js
const app = getApp()

Page({
  data: {
    userInfo: null,
    createTimeStr: '',
    lastLoginTimeStr: '',
    displayAvatarUrl: '', // 显示用的头像URL（可能是云存储下载后的临时路径）
    // 弹窗相关
    showEditModal: false,
    newAvatarUrl: '',
    newAvatarTempPath: '', // 本地临时路径
    newNickName: '',
    nickNameLength: 0,
    isUploading: false,
    uploadProgress: 0
  },

  onLoad(options) {
    this.loadUserInfo()
  },

  onShow() {
    // 页面显示时重新加载，以防从其他页面返回
    this.loadUserInfo()
  },

  // 加载用户信息（优先从缓存）
  loadUserInfo() {
    console.log('=== 开始加载用户信息 ===')
    
    // 1. 优先从本地缓存获取
    const cachedUserInfo = wx.getStorageSync('userInfo')
    
    if (cachedUserInfo) {
      console.log('从缓存加载用户信息:', cachedUserInfo)
      
      // 格式化时间
      const createTimeStr = this.formatTime(cachedUserInfo.createTime)
      const lastLoginTimeStr = this.formatTime(cachedUserInfo.lastLoginTime)

      this.setData({
        userInfo: cachedUserInfo,
        createTimeStr,
        lastLoginTimeStr,
        newNickName: cachedUserInfo.nickName || '',
        nickNameLength: (cachedUserInfo.nickName || '').length
      })

      // 加载云头像（如果是云存储路径）
      this.loadCloudAvatar(cachedUserInfo.avatarUrl)
      
      console.log('缓存加载完成')
    } else {
      console.log('缓存中没有用户信息，从服务器获取')
      
      // 2. 缓存没有，从服务器获取
      this.loadUserInfoFromCloud()
    }
  },

  // 从云数据库加载用户信息
  loadUserInfoFromCloud() {
    const app = getApp()
    
    wx.cloud.callFunction({
      name: 'getUserInfo',  // 需要创建专门的获取用户信息云函数
      data: {
        openId: app.globalData.openId
      }
    }).then(res => {
      console.log('从服务器获取用户信息:', res)
      
      if (res.result && res.result.success) {
        const userInfo = res.result.userInfo
        
        // 确保时间字段处理正确
        if (userInfo.createTime && typeof userInfo.createTime === 'object') {
          if (userInfo.createTime.$date) {
            userInfo.createTime = userInfo.createTime.$date
          }
        }
        if (userInfo.lastLoginTime && typeof userInfo.lastLoginTime === 'object') {
          if (userInfo.lastLoginTime.$date) {
            userInfo.lastLoginTime = userInfo.lastLoginTime.$date
          }
        }
        
        // 保存到本地缓存
        wx.setStorageSync('userInfo', userInfo)
        console.log('用户信息已保存到缓存')
        
        // 格式化时间
        const createTimeStr = this.formatTime(userInfo.createTime)
        const lastLoginTimeStr = this.formatTime(userInfo.lastLoginTime)

        // 更新页面
        this.setData({
          userInfo: userInfo,
          createTimeStr,
          lastLoginTimeStr,
          newNickName: userInfo.nickName || '',
          nickNameLength: (userInfo.nickName || '').length
        })

        // 加载云头像
        this.loadCloudAvatar(userInfo.avatarUrl)
        
        console.log('服务器加载完成')
      } else {
        console.error('获取用户信息失败:', res.result)
        // 返回登录页
        wx.reLaunch({
          url: '/pages/login/login'
        })
      }
    }).catch(err => {
      console.error('调用云函数失败:', err)
      // 返回登录页
      wx.reLaunch({
        url: '/pages/login/login'
      })
    })
  },

  // 加载云头像
  loadCloudAvatar(avatarUrl) {
    if (!avatarUrl) {
      this.setData({ displayAvatarUrl: '' })
      return
    }

    // 判断是否是云存储路径
    if (avatarUrl.startsWith('cloud://')) {
      console.log('检测到云存储头像，开始下载:', avatarUrl)
      
      // 下载云文件到本地临时路径
      wx.cloud.downloadFile({
        fileID: avatarUrl,
        success: (res) => {
          console.log('云头像下载成功:', res)
          this.setData({
            displayAvatarUrl: res.tempFilePath
          })
        },
        fail: (err) => {
          console.error('云头像下载失败:', err)
          // 下载失败时尝试使用原URL
          this.setData({
            displayAvatarUrl: avatarUrl
          })
        }
      })
    } else {
      // 普通URL直接使用
      this.setData({
        displayAvatarUrl: avatarUrl
      })
    }
  },

  // 格式化时间
  formatTime(timestamp) {
    if (!timestamp) return '未知'
    
    let date
    if (timestamp.$date) {
      // 处理云数据库返回的时间格式
      date = new Date(timestamp.$date)
    } else {
      date = new Date(timestamp)
    }
    
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    const hour = String(date.getHours()).padStart(2, '0')
    const minute = String(date.getMinutes()).padStart(2, '0')

    return `${year}-${month}-${day} ${hour}:${minute}`
  },

  // 打开编辑弹窗
  openEditModal() {
    this.setData({
      showEditModal: true,
      newAvatarUrl: '',
      newAvatarTempPath: '',
      newNickName: this.data.userInfo.nickName || '',
      nickNameLength: (this.data.userInfo.nickName || '').length,
      uploadProgress: 0
    })
  },

  // 关闭编辑弹窗
  closeEditModal() {
    if (this.data.isUploading) {
      wx.showToast({
        title: '上传中，请稍候',
        icon: 'none'
      })
      return
    }
    
    this.setData({
      showEditModal: false,
      newAvatarUrl: '',
      newAvatarTempPath: '',
      isUploading: false,
      uploadProgress: 0
    })
  },

  // 阻止事件冒泡
  stopPropagation() {
    // 阻止点击弹窗内容关闭弹窗
  },

  // 阻止touchmove事件
  preventTouchMove() {
    return false
  },

  // 选择头像（微信官方API）
  onChooseAvatar(e) {
    console.log('选择头像', e)
    const { avatarUrl } = e.detail
    
    if (avatarUrl) {
      // 本地预览
      this.setData({
        newAvatarUrl: avatarUrl,
        newAvatarTempPath: avatarUrl
      })
      
      wx.showToast({
        title: '头像已选择',
        icon: 'success',
        duration: 1500
      })
    }
  },

  // 昵称输入
  onNicknameInput(e) {
    const value = e.detail.value
    this.setData({
      newNickName: value,
      nickNameLength: value.length
    })
  },

  // 确认修改
  confirmEdit() {
    const { newNickName, newAvatarTempPath, userInfo } = this.data

    // 验证昵称
    if (!newNickName || newNickName.trim() === '') {
      wx.showToast({
        title: '请输入昵称',
        icon: 'none'
      })
      return
    }

    // 检查是否有修改
    const hasNickNameChange = newNickName.trim() !== (userInfo.nickName || '')
    const hasAvatarChange = newAvatarTempPath !== ''

    if (!hasNickNameChange && !hasAvatarChange) {
      wx.showToast({
        title: '未做任何修改',
        icon: 'none'
      })
      return
    }

    // 开始更新流程
    this.setData({ isUploading: true, uploadProgress: 0 })

    // 如果有新头像，先上传到云存储
    if (hasAvatarChange) {
      this.uploadAvatarAndUpdate(newAvatarTempPath, newNickName.trim())
    } else {
      // 没有新头像，只更新昵称
      this.updateUserInfoToCloud(newNickName.trim(), null)
    }
  },

  // 上传头像到云存储
  uploadAvatarAndUpdate(tempFilePath, nickName) {
    const openId = app.globalData.openId || 'unknown'
    const timestamp = Date.now()
    const cloudPath = `avatars/${openId}_${timestamp}.jpg`

    console.log('开始上传头像到云存储:', cloudPath)

    // 上传到云存储
    const uploadTask = wx.cloud.uploadFile({
      cloudPath: cloudPath,
      filePath: tempFilePath,
      success: (uploadRes) => {
        console.log('头像上传成功', uploadRes)
        const avatarUrl = uploadRes.fileID
        
        // 更新用户信息到数据库
        this.updateUserInfoToCloud(nickName, avatarUrl)
      },
      fail: (err) => {
        console.error('头像上传失败', err)
        this.setData({ isUploading: false, uploadProgress: 0 })
        
        wx.showModal({
          title: '上传失败',
          content: '头像上传失败，是否只更新昵称？',
          confirmText: '只更新昵称',
          cancelText: '取消',
          success: (res) => {
            if (res.confirm) {
              // 只更新昵称
              this.setData({ isUploading: true })
              this.updateUserInfoToCloud(nickName, null)
            }
          }
        })
      }
    })

    // 监听上传进度
    uploadTask.onProgressUpdate((res) => {
      console.log('上传进度:', res.progress)
      this.setData({
        uploadProgress: res.progress
      })
    })
  },

  // 调用云函数更新用户信息
  updateUserInfoToCloud(nickName, avatarUrl) {
    const updateData = {
      openId: app.globalData.openId
    }

    // 检查并处理昵称：去除首尾空白，确保不是空字符串
    if (nickName && nickName.trim() !== '') {
      updateData.nickName = nickName.trim()
    }

    if (avatarUrl) {
      updateData.avatarUrl = avatarUrl
    }

    console.log('调用云函数更新用户信息:', updateData)

    wx.cloud.callFunction({
      name: 'updateUserInfo',
      data: updateData
    }).then(res => {
      console.log('云函数返回结果:', res)
      
      if (res.result && res.result.success) {
        console.log('更新成功，准备重新加载用户信息')
        
        // 清空本地缓存
        wx.removeStorageSync('userInfo')
        console.log('已清除本地缓存')
        
        // 更新上传状态
        this.setData({
          isUploading: false,
          uploadProgress: 0
        })
        
        // 关闭弹窗
        this.closeEditModal()
        
        wx.showToast({
          title: '更新成功',
          icon: 'success'
        })
        
        // 延迟一下再重新加载，确保云数据库更新完成
        setTimeout(() => {
          this.loadUserInfoFromCloud()
        }, 500)
      } else {
        console.error('更新失败:', res.result)
        this.setData({ isUploading: false, uploadProgress: 0 })
        wx.showToast({
          title: res.result?.message || '更新失败',
          icon: 'none'
        })
      }
    }).catch(err => {
      console.error('调用云函数失败', err)
      this.setData({ isUploading: false, uploadProgress: 0 })
      
      // 开发阶段提示
      wx.showModal({
        title: '开发提示',
        content: '云函数未配置或调用失败\n是否使用本地模拟更新？',
        confirmText: '本地更新',
        cancelText: '取消',
        success: (modalRes) => {
          if (modalRes.confirm) {
            // 本地模拟更新（开发阶段）
            const updatedUserInfo = {
              ...this.data.userInfo,
              nickName: nickName,
              avatarUrl: avatarUrl || this.data.userInfo.avatarUrl,
              lastLoginTime: new Date().getTime()
            }
            wx.setStorageSync('userInfo', updatedUserInfo)
            
            this.setData({
              userInfo: updatedUserInfo,
              isUploading: false,
              createTimeStr: this.formatTime(updatedUserInfo.createTime),
              lastLoginTimeStr: this.formatTime(updatedUserInfo.lastLoginTime)
            })
            
            this.closeEditModal()
            
            wx.showToast({
              title: '本地更新成功',
              icon: 'success'
            })
          }
        }
      })
    })
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
