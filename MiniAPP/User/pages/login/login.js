// pages/login/login.js
const app = getApp()
Page({
  data: {
    agreementChecked: false,
    userInfo_tank: false,
    isUserExist: false,
    userInfo: {
      nickName: '',
      avatarURL: '',
    },
  },
  async onLoad() {
    try {
      await this.getOpenID()
    } catch (error) {
      console.error("获取openid失败", error);
    }
    this.getUserInfo()
  },

  // 处理登录点击
  handleLoginClick: function () {
    if (!this.data.agreementChecked) {
      wx.showToast({
        title: '请先同意用户协议',
        icon: 'none',
        duration: 2000
      });
      return
    }
    // 弹窗控制
    this.TankControl()
    // 用户信息已存在则免注册直接登录
    if (this.data.isUserExist) {
      this.wxLogin()
    }
  },

  //关闭/打开弹窗
  TankControl(e) {
    if (!this.data.userInfo_tank) {
      //检查用户是否存在
      if (this.data.isUserExist) {
        this.setData({
          userInfo_tank: false
        })
      } else {
        this.setData({
          userInfo_tank: true
        })
      }
    } else {
      this.setData({
        userInfo_tank: false
      })
    }

  },
  getUserInfo() {
    let userInfo = wx.getStorageSync('userInfo')
    if (userInfo) {
      this.setData({
        userInfo,
        isUserExist: true
      })
      console.log("缓存中获取到用户信息",userInfo)
      return
    }
    console.log("获取用户信息的openid",app.globalData.openid)
    wx.cloud.callFunction({
      name: 'getUserInfo',
      data: {
        openid: app.globalData.openid,
      }
    }).then(res => {
      if (res.result && res.result.success) {
        this.setData({
          'userInfo.nickName': res.result.data.nickName,
          'userInfo.avatarURL': res.result.data.avatarURL,
          isUserExist: true
        })
        //缓存用户信息
        wx.setStorageSync('userInfo', this.data.userInfo)
        console.log(res.result)
      } else {
        this.setData({
          isUserExist: false
        })
        console.log("获取用户信息失败",res.result)
      }
    }).catch(err => {
      console.error("调用云函数 getUserInfo 失败", err); // 捕获云函数调用本身的错误
    })
  },
  getOpenID(){
    return new Promise((resolve, reject) => {
      const openid = wx.getStorageSync('openid')
      if (openid) {
        app.globalData.openid=openid
        resolve(openid)
        return
      }
      wx.cloud.callFunction({
        name: 'getOpenID',
        data: {}
      }).then(res => {
        // 检查云函数返回的实际结果中的 success 字段
        if (res.result && res.result.success) {
          // 直接从 res.result 中获取 openid
          const openid = res.result.openid;
          // 更新 globalData
          app.globalData.openid = openid;
          resolve(openid)
          console.log("获取openid成功", openid); 
          // 缓存 openid
          wx.setStorageSync('openid', openid);

        } else {
          console.log("获取openid失败", res); // 打印完整的 res 对象以便调试
          reject(res)
        }
      }).catch(err => {
        console.error("调用云函数 getOpenID 失败", err); // 捕获云函数调用本身的错误
        reject(err) 
      });
    })
    },
  //获取头像
  onChooseAvatar(e) {
    console.log(e.detail);
    this.setData({
      'userInfo.avatarURL': e.detail.avatarUrl
    })
  },
  //获取用户昵称
  getNickName(e) {
    console.log(e.detail)
    this.setData({
      'userInfo.nickName': e.detail.value
    })
  },
  // 提交
  async submit(e) {
    // 限制注册用户信息非空
    if (!this.data.userInfo.avatarURL) {
      return wx.showToast({
        title: '请选择头像',
        icon: 'error'
      })
    }
    if (!this.data.userInfo.nickName) {
      return wx.showToast({
        title: '请输入昵称',
        icon: 'error'
      })
    }
    this.setData({
      userInfo_tank: false
    })
    console.log("=====正在注册用户信息中====")
    //上传头像
    try {
      await this.uploadAvatarFile()
    } catch (error) {
      console.error("上传头像失败", error);
    }
    // 登录
    this.wxLogin()

  },
  uploadAvatarFile() {
    return new Promise((resolve, reject) => {
      //上传临时头像图片到云存储
      let tempPath = this.data.userInfo.avatarURL;

      let suffix = /\.[^\.]+$/.exec(tempPath)[0];//正则表达提取文件扩展名
      console.log("====上传头像中====");

      wx.cloud.uploadFile({
        cloudPath: 'userImg/' + this.data.userInfo.nickName +'/' + new Date().getTime() + suffix, //在云端的文件名称
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
  // 切换协议同意状态
  toggleAgreement: function () {
    this.setData({
      agreementChecked: !this.data.agreementChecked
    });
  },

  // 查看协议详情
  navigateToAgreement: function (e) {
    const type = e.currentTarget.dataset.type;
    const title = type === 'user' ? '用户协议' : '隐私政策';

    wx.showModal({
      title: title,
      content: '即将跳转到' + title + '页面',
      confirmText: '前往',
      success: (res) => {
        if (res.confirm) {
          const url = type === 'user'
            ? '/pages/agreement/user/user'
            : '/pages/agreement/privacy/privacy';

          wx.navigateTo({
            url: url
          });
        }
      }
    });
  },


  wxLogin: function () {
    console.log("当前登录的用户信息为:", app.globalData.openid,this.data.userInfo)
    wx.cloud.callFunction({
      name: 'login',
      data: {
        openid: app.globalData.openid,
        nickName: this.data.userInfo.nickName,
        avatarURL: this.data.userInfo.avatarURL,
      }
    }).then(res => {
      if (res.result && res.result.success) {
        console.log("登录成功,返回数据为:", res.result)
      } else {
        console.log("登录失败,返回数据为:", res.result)
      }
    }).catch(err => {
      console.error("云函数调用失败:", err);
    })
    wx.navigateTo({
      url: '/pages/control-panel/control',
    })
  },
});
