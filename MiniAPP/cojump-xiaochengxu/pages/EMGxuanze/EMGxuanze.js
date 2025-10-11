// index.js
const app = getApp();

Page({
  data: {
    isConnected: false,
    connecting: false
  },

  stringToBytes(str) {
    var array = new Uint8Array(str.length);
    for (var i = 0, l = str.length; i < l; i++) {
      array[i] = str.charCodeAt(i);
    }
    console.log(array);
    return array.buffer;
  },
  
  onShow() {
    this.setData({
      isConnected: app.globalData.isConnected
    });
  },
  
  connectDevice() {
    if (this.data.connecting || this.data.isConnected) return;
    
    this.setData({ connecting: true });
    wx.showLoading({ title: '连接设备中...', mask: true });
    
    app.initBluetooth()
      .then(() => {
        console.log('蓝牙适配器初始化成功');
        return app.searchDevice(app.globalData.deviceName);
      })
      .then(device => {
        console.log('找到设备:', device);
        return app.connectDevice(device);
      })
      .then(() => {
        console.log('设备连接成功');
        this.setData({ 
          isConnected: true,
          connecting: false
        });
        wx.showToast({ title: '连接成功', icon: 'success' });
      })
      .catch(err => {
        console.error('连接失败:', err);
        wx.showToast({ 
          title: '连接失败: ' + (err.message || err), 
          icon: 'none',
          duration: 3000
        });
        this.setData({ connecting: false });
      })
      .finally(() => {
        wx.hideLoading();
      });
  },
  
  navigateToCeliang() {
    /*if (!this.data.isConnected) {
      wx.showToast({ title: '请先连接设备', icon: 'none' });
      return;
    }*/
    wx.navigateTo({ url: '/pages/EMGceliang/EMGceliang' });
  },
  
  navigateToTrain() {
    /*if (!this.data.isConnected) {
      wx.showToast({ title: '请先连接设备', icon: 'none' });
      return;
    }*/
    wx.navigateTo({ url: '/pages/EMGtrain/EMGtrain' });
  },


  

  

});