// index.js
/*
  肌力训练页面
*/
// 获取应用实例
const app = getApp()

Page({
  data: {
    'speedone': '',
    'speedtwo': '',
    'voltageValue': '0.00', // 用来存储电压信息
    'speedonepass':'',
    'speedtwopass':'',
    inputTime: '', // 绑定到输入框的初始空值
    'time': '', // 默认倒计时时间
    timeStr: '', // 倒计时展示的字符串
    timer: null, // 定时器
    timeInSecs: 0,
    timerStatus:"pause",
  },
  
  stringToBytes(str) {
    var array = new Uint8Array(str.length);
    for (var i = 0, l = str.length; i < l; i++) {
      array[i] = str.charCodeAt(i);
    }
    console.log(array);
    return array.buffer;
  },
  hextoString: function (hex) {
    var arr = hex.split("")
    var out = ""
    for (var i = 0; i < arr.length / 2; i++) {
      var tmp = "0x" + arr[i * 2] + arr[i * 2 + 1]
      var charValue = String.fromCharCode(tmp);
      out += charValue
    }
    this.setData({
      voltageValue: out,
    });
    return out
  },
  ab2hex(buffer) {
    var hexArr = Array.prototype.map.call(
      new Uint8Array(buffer),
      function (bit) {
        return ('00' + bit.toString(16)).slice(-2)
      }
    )
    return hexArr.join('');
  },
  starting(){
    var buffer = this.stringToBytes("start")
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
    console.log('开始')
  },
  pausing(){
    var buffer = this.stringToBytes("pause")
    console.log('暂停')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  },
  mode1(){
    var buffer = this.stringToBytes("mode1")
    console.log('模式1')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  },
  mode2(){
    var buffer = this.stringToBytes("mode2")
    console.log('模式2')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  },
  mode3(){
    var buffer = this.stringToBytes("mode3")
    console.log('模式3')
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  },
// 将数字转换为字节流的函数
  numberToBytes(num) {
    // 确保输入在有效范围内
  // if (num < 0 || num > 999) {
  //   console.error('输入的数字超出范围，请输入0-999之间的整数');
  //   return null;
  // }

  // 创建一个 2 字节的 ArrayBuffer
  const buffer = new ArrayBuffer(2);
  const dataView = new DataView(buffer);

  // 将 num 转换为 16 位整数
  const int16Value = num & 0xFFFF; // 限制在 16 位范围内

  // 使用小端序存储
  dataView.setUint16(0, int16Value, true);

  return buffer;
  },
  // 输入框输入时触发
  bindInput: function(e) {
    const inputTime = parseInt(e.detail.value, 10); // 将输入值转换为整数
    if (!isNaN(inputTime) && inputTime >= 0) {
      this.setData({
        inputTime: e.detail.value,
        timeInSecs: inputTime * 60, // 将输入的分钟数转换为秒数
        timeStr: this.formatTime(inputTime * 60), // 格式化时间字符串
      });
    }
  },
  // 格式化时间显示为 MM:SS
  formatTime: function(seconds) {
    const minutes = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  },
  start: function() {
    this.startTimer();
    this.starting();
  },
  pause: function() {
    this.pauseTimer();
    this.pausing();
  },
  // 开始倒计时
  startTimer: function() {
    const that = this;
    that.setData({
      inputTime: '', // 清空输入框
    });

    if (that.data.timerStatus == "start") {
      return;
    }
    that.data.timer = setInterval(function() {
      if (that.data.timeInSecs > 0) {
        that.setData({
          timeInSecs: that.data.timeInSecs - 1,
          timeStr: that.formatTime(that.data.timeInSecs), // 每次减少秒数时更新时间字符串
        });
      } else {
        that.endTimer(); // 结束倒计时
      }
    }, 1000);
    that.setData({
      timerStatus:"start",
    });
  },
  // 暂停倒计时
  pauseTimer: function() {
    const that = this;
    if (that.data.timer) {
      clearInterval(that.data.timer);
      that.setData({
        timerStatus:"pause",
      });
    }
  },
  // 结束倒计时
  endTimer: function() {
    const that = this;
    if (that.data.timer) {
      clearInterval(that.data.timer);
      that.setData({
        timeInSecs: 0,
        timeStr: '00:00', // 设置结束时间字符串
      });
    }
    this.pausing()
  },
  // 页面卸载时清除定时器
  onUnload: function() {
    if (this.data.timer) {
      clearInterval(this.data.timer);
    }
  },
})