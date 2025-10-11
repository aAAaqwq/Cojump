// app.js
App({
  onLaunch() {
    // 展示本地存储能力
    //修改
    const logs = wx.getStorageSync('logs') || []
    logs.unshift(Date.now())
    wx.setStorageSync('logs', logs)
    wx.app = this;

    //初始化云函数
    wx.cloud.init({
      //云环境id
      env: "cloudbase-7gky5kj9737551d9"
    })

    //获取openid
    wx.cloud.callFunction({
      name: 'getOpenID',
      success: res => {
        this.globalData.openId = res.result.openid;
        console.log('openid:', this.globalData.openId);
      },
      fail: err => {
        console.error('获取openid失败:', err);
      }
    });
    
      
    

    // 登录
    wx.login({
      success: () => {
        // 发送 res.code 到后台换取 openId, sessionKey, unionId
      }
    })

    // 获取用户信息
    wx.getSetting({
      success: res => {
        if (res.authSetting['scope.userInfo']) {
          // 已经授权，可以直接调用 getUserInfo 获取头像昵称，不会弹框
          wx.getUserInfo({
            success: res => {
              // 可以将 res 发送给后台解码出 unionId
              this.globalData.userInfo = res.userInfo

              // 由于 getUserInfo 是网络请求，可能会在 Page.onLoad 之后才返回
              // 所以此处加入 callback 以防止这种情况
              if (this.userInfoReadyCallback) {
                this.userInfoReadyCallback(res)
              }
            }
          })
        }
      }
    })
  },
  // app.js
globalData: {
  openId:"",
  userInfo: null,
  'deviceId':'',
  'serviceId':'',
  'characteristicId':'',
  'latestBleData': '',
  bleDataHistory: [],
  latestBleData: '',
  eventBus: {  // 独立的事件总线对象
    _events: {},
    $on(name, cb) {
      this._events[name] = this._events[name] || [];
      this._events[name].push(cb);
    },
    $emit(name, ...args) {
      this._events[name]?.forEach(cb => cb(...args));
    },
    $off(name, cb) {
      if (!this._events[name]) return;
      
      // 如果没有提供回调函数，移除该事件的所有监听器
      if (!cb) {
        this._events[name] = [];
        return;
      }
      
      // 移除特定的回调函数
      this._events[name] = this._events[name].filter(callback => callback !== cb);
    }
  }
},

  // app.js
sendCommand(command) {
  return new Promise((resolve, reject) => {
    if (!this.globalData.deviceId || !this.globalData.serviceId || !this.globalData.characteristicId) {
      reject('蓝牙未连接或连接参数不完整');
      return;
    }
    
    const buffer = new ArrayBuffer(command.length);
    const dataView = new Uint8Array(buffer);
    
    for (let i = 0; i < command.length; i++) {
      dataView[i] = command.charCodeAt(i);
    }
    
    wx.writeBLECharacteristicValue({
      deviceId: this.globalData.deviceId,
      serviceId: this.globalData.serviceId,
      characteristicId: this.globalData.characteristicId,
      value: buffer,
      success: resolve,
      fail: reject
    });
  });
}

  
})

