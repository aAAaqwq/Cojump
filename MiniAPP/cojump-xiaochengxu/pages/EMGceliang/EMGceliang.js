// pages/page3/page3.js
const app = getApp()

Page({
  data: {
    recommendAvg: 0, // 新增：推荐阈值
    // 校准状态文本
    stepText: ['准备校准', '请放松肌肉', '请用力收缩', '校准完成'],
    // 测量状态相关
    isMeasuring: false,
    currentStep: 0,
    relaxTime: 10,
    contractTime: 5,
    relaxAvg: 0,
    contractAvg: 0,
    
    // 蓝牙状态相关
    isConnected: false,
    deviceId: '',
    serviceId: '',
    characteristicId: '',
    
    // 倒计时定时器
    countdownTimer: null
  },

  onLoad() {
    // wx.setStorageSync('emgThreshold', 360)
    // console.log("设置成功")
    // 获取全局蓝牙连接信息
    this.setData({
      deviceId: app.globalData.deviceId,
      serviceId: app.globalData.serviceId,
      characteristicId: app.globalData.characteristicId,
      isConnected: !!app.globalData.deviceId
    })

    // 监听蓝牙数据更新
    const that = this;
    app.globalData.eventBus.$on('bleDataUpdate', (data) => {
      that.handleBLEData(data);
    });

    // 如果未连接蓝牙，提示返回首页连接
    if (!this.data.isConnected) {
      wx.showModal({
        title: '未连接设备',
        content: '请先返回首页连接蓝牙设备',
        showCancel: false,
        confirmText: '返回',
        success: () => {
          wx.navigateBack();
        }
      });
    }
  },

  onUnload() {
    // 移除事件监听并清除定时器
    app.globalData.eventBus.$off('bleDataUpdate');
    this.clearCountdown();
  },

  // 处理蓝牙接收的数据
 // 处理蓝牙接收的数据

 handleBLEData(data) {
  console.log('原始数据:', data); // 打印完整接收到的数据
  
  // 处理校准阶段提示
  if (data === 'CALIB_START:RELAX') {
    if (this.data.currentStep !== 1) { 
      this.clearCountdown();
      this.setData({
        currentStep: 1,
        relaxTime: 10
      });
      this.startRelaxCountdown();
    }
  } else if (data === 'CALIB_STAGE:CONTRACT') {
    if (this.data.currentStep !== 2) { 
      this.clearCountdown();
      this.setData({
        currentStep: 2,
        contractTime: 5
      });
      this.startContractCountdown();
    }
  } 
  // 处理校准进度（倒计时更新）
  else if (data.startsWith('RELAX,')) {
    const remaining = parseInt(data.split(',')[1]);
    this.setData({ relaxTime: remaining });
  } else if (data.startsWith('CONTRACT,')) {
    const remaining = parseInt(data.split(',')[1]);
    this.setData({ contractTime: remaining });
  }
  // 处理校准结果（分字段接收）
  else if (data.startsWith('Relax=')) {
    this.setData({ relaxAvg: parseFloat(data.split('=')[1]) });
  } else if (data.startsWith('Contract=')) {
    this.setData({ contractAvg: parseFloat(data.split('=')[1]) });
  } else if (data.startsWith('Recommend=')) {
    this.setData({ 
      recommendAvg: parseFloat(data.split('=')[1]),
      currentStep: 3,
      isMeasuring: false
    });
  }
  // 处理模式切换反馈
  else if (data === 'EMG_MODE_ACTIVE') {
    wx.showToast({ title: '已切换到EMG模式', icon: 'none' });
  } else if (data === 'BLUETOOTH_MODE_ACTIVE') {
    wx.showToast({ title: '已切换到蓝牙模式', icon: 'none' });
  }
},

  // 开始测量
  startMeasurement() {
    if (!this.data.isConnected) {
      wx.showToast({
        title: '未连接设备',
        icon: 'none'
      });
      return;
    }
    // 使用 app.js 中的 sendCommand 方法发送命令
    app.sendCommand('emgcalibrate')
      .then(() => {
        // 初始化测量状态并启动倒计时
        this.setData({
          isMeasuring: true,
          currentStep: 0, // 新增：等待设备响应
          relaxAvg: 0,
          contractAvg: 0,
          recommendAvg: 0
        });
        wx.showToast({
          title: '校准已开始',
          icon: 'none',
          duration: 1500
        });
      })
      .catch(err => {
        console.error('发送命令失败:', err);
        wx.showToast({
          title: '发送命令失败',
          icon: 'none'
        });
      });
  },

  startRelaxCountdown() {
    // 先清除可能存在的旧定时器
    this.clearCountdown();
    
    const timer = setInterval(() => {
      const currentTime = this.data.relaxTime - 1;
      
      if (currentTime >= 0) {
        this.setData({ relaxTime: currentTime });
      } else {
        clearInterval(timer);
        this.setData({
          currentStep: 2,
          contractTime: 5
        });
        this.startContractCountdown();
      }
    }, 1000);
    
    // 用变量缓存定时器，而非存在data中（避免setData延迟）
    this.countdownTimer = timer; 
  },
  
  // 同步修改 startContractCountdown 方法
  startContractCountdown() {
    this.clearCountdown();
    
    const timer = setInterval(() => {
      const currentTime = this.data.contractTime - 1;
      
      if (currentTime >= 0) {
        this.setData({ contractTime: currentTime });
      } else {
        clearInterval(timer);
        this.setData({
          currentStep: 3,
          isMeasuring: false
        });
      }
    }, 1000);
    
    this.countdownTimer = timer; // 用变量缓存
  },
  
  // 修改清除定时器方法
  clearCountdown() {
    if (this.countdownTimer) { // 直接访问变量，而非data
      clearInterval(this.countdownTimer);
      this.countdownTimer = null;
    }
  },

  // 应用推荐阈值
  applyRecommendedThreshold() {
    if (!this.data.isConnected) return;
    
    app.sendCommand('applyrecommended')
      .then(() => {
        wx.showToast({
          title: '阈值已应用',
          icon: 'success'
        });
      })
      .catch(err => {
        console.error('发送命令失败:', err);
        wx.showToast({
          title: '发送命令失败',
          icon: 'none'
        });
      });
      // 记录阈值
      wx.setStorageSync('emgThreshold', this.data.recommendAvg)
      // 数据存储到云端
      
  },

  // 返回上一页
  navigateBack() {
    this.clearCountdown();
    wx.navigateBack();
  }
})  