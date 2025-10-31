// train.js
const app = getApp();

Page({
  data: {
    emgThreshold: 30,
    isTraining: false,
    statusMessage: '等待开始',
    triggerCount: 0,
    currentSignal: 0
  },
  
  onLoad() {
    const tempThreshold=wx.getStorageSync("emgThreshold")
    if (tempThreshold){
      this.setData({
        emgThreshold:tempThreshold
      }
      )
    }
    // 设置数据接收回调
    app.globalData.onDataReceived = this.handleBLEData.bind(this);
  },

  stringToBytes(str) {
    var array = new Uint8Array(str.length);
    for (var i = 0, l = str.length; i < l; i++) {
      array[i] = str.charCodeAt(i);
    }
    console.log(array);
    return array.buffer;
  },
  
  // 格式化日期为 YYYY-MM-DD HH:MM:SS 格式
  formatDate(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
  },
  
  onThresholdInput(e) {
    this.setData({ emgThreshold: e.detail.value });
  },
  
  setThreshold() {
    // 获取输入的阈值并转换为数值
    const threshold = parseFloat(this.data.emgThreshold);
    
    if (!isNaN(threshold)) {
        // 构建T45格式的字符串
        const dataString = `T${threshold}`;
        // 将字符串转换为ArrayBuffer
        const buffer = this.stringToArrayBuffer(dataString);
        
        wx.writeBLECharacteristicValue({
            deviceId: app.globalData.deviceId,
            serviceId: app.globalData.serviceId,
            characteristicId: app.globalData.characteristicId,
            value: buffer,
            success: (res) => {
                console.log('阈值发送成功');
                wx.showToast({
                    title: '阈值发送成功',
                    icon: 'success'
                });
            },
            fail: (err) => {
                console.error('阈值发送失败:', err);
                wx.showToast({
                    title: '阈值发送失败',
                    icon: 'none'
                });
            }
        });

        this.storeEMGData()
    } else {
        wx.showToast({
            title: '请输入有效的阈值',
            icon: 'none'
        });
    }
},

storeEMGData(){
  wx.cloud.callFunction({
    name: 'setEMG',
    data: {
      deviceId: app.globalData.deviceId,
      openId: app.globalData.openId,
      emgThreshold: this.data.emgThreshold,
      emgRaw:[],
      recordTime: this.formatDate(new Date()),
    }
  }).then(res => {
    console.log('数据存储成功:', res);
    wx.showToast({
      title: '数据存储成功',
      icon: 'success'
    });
  }).catch(err => {
    console.error('数据存储失败:', err);
    wx.showToast({
      title: '数据存储失败',
      icon: 'none'
    });
  });
},

// 将字符串转换为ArrayBuffer的函数
stringToArrayBuffer(str) {
    const buffer = new ArrayBuffer(str.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i < str.length; i++) {
        view[i] = str.charCodeAt(i);
    }
    return buffer;
},

start() {
    this.setThreshold();
},
  
  
  // 其他方法保持不变...
  emgmode(){
    var buffer = this.stringToBytes("emgmode")
    console.log('切换至emg模式')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
    wx.showToast({
      title: '已切换至EMG模式',
      icon: 'success',
      duration: 1600
    });
    this.setData({
      statusMessage: '已切换至EMG模式',
      isTraining: true
    })
  },
  bluetoothmode(){
    var buffer = this.stringToBytes("bluetoothmode")
    console.log('切换回蓝牙模式')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
    wx.showToast({
      title: '切换回蓝牙模式',
      icon: 'success',
      duration: 2000 // 显示时间，单位ms
    });
  },

  

  onUnload() {
    this.bluetoothmode();
  },

  ToPage1: function() {
    wx.navigateTo({
      url: '/pages/EMGxuanze/EMGxuanze'
    });
  console.log('跳转至emg选择界面')
  },

});