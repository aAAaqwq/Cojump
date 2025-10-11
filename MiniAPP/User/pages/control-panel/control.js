const app = getApp();
Page({
  data: {
    userInfo: {
      avatarUrl: '/images/default-avatar.png',
      nickName: '微信用户'
    },
    tempAvatarUrl: '',
    tempNickname: '',
    showModal: false,
    showToast: false,
    showEMGDataTable: false, // 控制EMG数据表格的显示
    showEMGInput: false,     // 控制EMG数据输入区域的显示
    emgData: [],             // 存储EMG数据
    emgInputArrayStr: '',    // 存储EMG输入字符串
    emgDeviceId: '',         // 存储设备ID
    emgThreshold: '',        // 存储阈值
  },

  async onShow() {
    // 尝试从本地存储加载用户信息
    try {
      await this.loadUserInfo();
    } catch (error) {
      console.error("加载用户信息失败", error);
    }
    console.log("加载用户信息为：", this.data.userInfo)
    // 加载头像                     
    try {
      await this.saveAvatarToLocal(this.data.userInfo.avatarURL);
    } catch (error) {
      console.error("加载头像失败", error);
    }
  },

  // 加载用户信息
  loadUserInfo: function () {
    return new Promise((resolve, reject) => {
      //缓存中获取
      const userInfo = wx.getStorageSync('userInfo');
      if (userInfo) {
        this.setData({ userInfo });
        resolve(userInfo);
      } else {
        // 如果没有本地信息，尝试从云端获取
        wx.cloud.callFunction({
          name: 'getUserInfo',
          data: {
            openid: app.globalData.openid,
          },
          success: (res) => {
            if (res.result.userInfo) {
              this.setData({ userInfo: res.result.userInfo });
              // 缓存
              wx.setStorageSync('userInfo', res.result.userInfo);
              resolve(res.result.userInfo);
            } else {
              reject(new Error('获取用户信息失败'));
            }
          }
        })
      }
    });
  },

  // 从云存储中获取头像文件到本地临时文件保存
  saveAvatarToLocal: function (avatarURL) {
    return new Promise(async (resolve, reject) => {
      console.log("====下载头像到temp中====", avatarURL)
      // 如果是云存储路径，直接使用 wx.cloud.downloadFile
      if (avatarURL && avatarURL.startsWith('cloud://')) {
        try {
          const res = await wx.cloud.downloadFile({
            fileID: avatarURL,
          });
          if (res.statusCode === 200 && res.tempFilePath) {
            //缓存临时路径到本地
            this.setData({
              tempAvatarUrl: res.tempFilePath
            });
            resolve(res.tempFilePath);
          } else {
            console.error('云文件下载失败', res);
            return reject(new Error(`云文件下载失败，状态码: ${res.statusCode}`));
          }
        } catch (error) {
          console.error('调用 wx.cloud.downloadFile 失败', error);
          return reject(error);
        }
      } else if (avatarURL) {
        // 如果是普通 HTTP/HTTPS URL，使用 wx.downloadFile
        wx.downloadFile({
          url: avatarURL,
          success: (res) => {
            if (res.statusCode === 200) {
              this.setData({
                tempAvatarUrl: res.tempFilePath
              })
              resolve(res.tempFilePath);
            } else {
              reject(new Error(`下载头像失败，状态码: ${res.statusCode}`));
            }
          },
          fail: (err) => {
            reject(err);
          }
        });
      } else {
        console.error('下载URL为空');
        return reject(new Error('下载URL为空'));
      }
    });
  },

  // 上传头像到云存储
  uploadAvatarFile() {
    return new Promise((resolve, reject) => {
      //上传临时头像图片到云存储
      let tempPath = this.data.tempAvatarUrl;
      if (!tempPath) {
        console.error('临时文件路径为空');
        return reject(new Error('临时文件路径为空'));
      }

      // 使用正则表达式提取文件扩展名，如果无法提取则默认为.png
      let suffix = '.png';
      try {
        const extMatch = /\.[^\.]+$/.exec(tempPath);
        if (extMatch && extMatch[0]) {
          suffix = extMatch[0];
        }
      } catch (error) {
        console.warn('无法提取文件扩展名，使用默认值', error);
      }
      
      console.log("====上传头像中====");
      
      // 使用tempNickname作为文件夹名，如果为空则使用默认值
      const folderName = this.data.tempNickname || 'default_user';

      wx.cloud.uploadFile({
        cloudPath: 'userImg/' + folderName + '/' + new Date().getTime() + suffix, //在云端的文件名称
        filePath: tempPath, // 临时文件路径
        success: res => {
          console.log('上传成功', res)
          let fileID = res.fileID
          //更新avatarURL
          this.setData({
            'userInfo.avatarURL': fileID,
          })
          resolve(fileID)
        },
        fail: err => {
          console.log('上传头像失败', err)
          wx.showToast({
            icon: 'error',
            title: '上传头像错误',
          })
          reject(err)
        }
      })
    })
  },
  //更新用户信息
  updateUserInfo: function (userInfo) {
    return new Promise((resolve, reject) => {
      // 更新缓存
      wx.setStorageSync('userInfo', userInfo);
      // 更新数据库
      wx.cloud.callFunction({
        name: 'updateUserInfo',
        data: userInfo,
        success: (res) => {
          if (res.result.success) {
            resolve(res.result);
          } else {
            reject(new Error(res.result.message));
          }
        },
        fail: (err) => {
          reject(err);
        }
      })
    })
  },

  // 显示编辑浮窗
  showEditModal: function () {
    this.setData({
      showModal: true,
      tempAvatarUrl: this.data.tempAvatarUrl,
      tempNickname: this.data.tempNickname
    });
  },

  // 隐藏浮窗
  hideModal: function () {
    this.setData({ showModal: false });
  },

  // 选择头像处理
  onChooseAvatar: function (e) {
    const { avatarUrl } = e.detail;
    this.setData({ tempAvatarUrl: avatarUrl });
  },

  // 昵称输入处理
  onNicknameInput: function (e) {
    this.setData({ tempNickname: e.detail.value });
  },

  // 保存个人信息
  async saveProfile() {
    const { tempAvatarUrl, tempNickname } = this.data;

    if (!tempNickname || tempNickname.trim() === '') {
      wx.showToast({
        title: '请输入昵称',
        icon: 'none'
      });
      return;
    }
    console.log("准备上传到云:", tempAvatarUrl, tempNickname)
    //上传更新到云
    try {
      //上传头像到云存储
      const fileID = await this.uploadAvatarFile()
      console.log("更新云数据库:", fileID)
      //更新用户信息到数据库
      const res = this.updateUserInfo({
        openid: app.globalData.openid,
        avatarURL: fileID, 
        nickName: tempNickname
      })
      console.log("更新用户信息结果:", res)
    } catch (error) {
      console.error('上传到云失败', error)
      wx.showToast({
        icon: 'error',
        title: '上传错误',
      })
    }
    this.setData({
      showModal: false
    });
    // 显示保存成功提示
    this.showSuccessToast();
    // 刷新页面
    this.onShow()
  },

  // 显示成功提示
  showSuccessToast: function () {
    this.setData({ showToast: true });

    setTimeout(() => {
      this.setData({ showToast: false });
    }, 2000);
  },


  // 查询EMG数据按钮点击事件
  onQueryEMGData: function () {
    this.setData({
      showEMGDataTable: !this.data.showEMGDataTable, // 切换显示状态
      showEMGInput: false, // 隐藏设置区域
    });
    // 模拟加载EMG数据
    // if (this.data.showEMGDataTable) {
    //   const mockData = [
    //     { time: '10:00:00', value: 1.23 },
    //     { time: '10:00:01', value: 2.45 },
    //     { time: '10:00:02', value: 3.67 },
    //     { time: '10:00:03', value: 4.89 },
    //   ];
    //   this.setData({ emgData: mockData });
    // }
    //加载EMG数据
    this.loadEMGData()
  },
  //加载EMG数据
  loadEMGData() {
    wx.cloud.callFunction({
      name: 'getEMG',
      data: {
        openid: app.globalData.openid,
      },
      success: (res) => {
        if (res.result.success) {
          console.log('查询EMG数据成功:', res.result.data)
          const emgData = res.result.data.map(item => ({
            timestamp: item.timestamp,
            emg_raw: item.emg_raw,
            threshold: item.threshold,
          }))
          this.setData({
            emgData: emgData,
          })
        } else {
          wx.showToast({
            icon: 'error',
            title: res.result.message,
          })
        }
      },
      fail: (err) => {
        wx.showToast({
          icon: 'error',
          title: '查询EMG数据失败',
        })
      }
    })
  },

  // 设置EMG数据按钮点击事件
  onSetEMGData: function () {
    this.setData({
      showEMGInput: !this.data.showEMGInput, // 切换显示状态
      showEMGDataTable: false, // 隐藏查询区域
      emgInputArrayStr: '', // 清空输入框
      emgDeviceId: app.globalData.dev_id || '', // 默认使用全局设备ID
      emgThreshold: '', // 清空阈值
    });
  },

  // 设备ID输入处理
  onDeviceIdInput: function (e) {
    this.setData({ emgDeviceId: e.detail.value });
  },

  // EMG数据输入处理
  onEMGInput: function (e) {
    this.setData({ emgInputArrayStr: e.detail.value });
  },

  // 阈值输入处理
  onThresholdInput: function (e) {
    this.setData({ emgThreshold: e.detail.value });
  },

  // 确认设置EMG数据
  onConfirmSetEMGData: function () {
    // 验证表单数据
    if (!this.data.emgDeviceId) {
      wx.showToast({
        title: '请输入设备ID',
        icon: 'none',
        duration: 2000
      });
      return;
    }

    if (!this.data.emgInputArrayStr) {
      wx.showToast({
        title: '请输入EMG原始值数组',
        icon: 'none',
        duration: 2000
      });
      return;
    }

    if (!this.data.emgThreshold) {
      wx.showToast({
        title: '请输入阈值',
        icon: 'none',
        duration: 2000
      });
      return;
    }

    // 处理EMG原始值数组
    const inputStr = this.data.emgInputArrayStr;
    const floatArray = inputStr.split(' ').map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
    
    if (floatArray.length === 0) {
      wx.showToast({
        title: 'EMG原始值数组格式错误',
        icon: 'none',
        duration: 2000
      });
      return;
    }

    // 处理阈值
    const threshold = parseFloat(this.data.emgThreshold);
    if (isNaN(threshold)) {
      wx.showToast({
        title: '阈值必须是有效数字',
        icon: 'none',
        duration: 2000
      });
      return;
    }

    console.log('确认设置EMG数据:', {
      deviceId: this.data.emgDeviceId,
      emgRaw: floatArray,
      threshold: threshold
    });

    // 上传数据到云端
    this.uploadEMG(this.data.emgDeviceId, floatArray, threshold);
    
    wx.showToast({
      title: 'EMG数据设置成功',
      icon: 'success',
      duration: 2000
    });
    this.setData({ showEMGInput: false }); // 隐藏设置区域
  },

  // 取消设置EMG数据
  onCancelSetEMGData: function () {
    console.log('取消设置EMG数据');
    this.setData({ showEMGInput: false }); // 隐藏设置区域
  },

  // 上传EMG数据到云端
  uploadEMG(deviceId, floatArray, threshold) {
    wx.cloud.callFunction({
      name: 'setEMG',
      data: {
        openid: app.globalData.openid,
        dev_id: deviceId,
        emg_raw: floatArray,
        threshold: threshold,
        timestamp: new Date().toLocaleString('zh-CN', {
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: false // 使用24小时制
        }),
      },
      success: (res) => {
        if (res.result.success) {
          console.log('上传EMG数据成功:', res.result.data);
        } else {
          console.log('上传EMG数据失败:', res.result.message);
        }
      },
      fail: (err) => {
        console.log('上传EMG数据失败:', err);
      }
    })
  }

});