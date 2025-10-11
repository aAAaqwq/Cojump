// index.js
/*
  热身模式页面
*/
// 获取应用实例
const app = getApp()
const util = require('../../utils/util.js')

Page({
  data: {
    speedInput: ''
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
  lightoff(){
    var buffer = this.stringToBytes("lightoff")
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  },
  setspeed() {
    // 获取输入的阈值并转换为数值
    const speed = parseFloat(this.data.speed);
    
    if (!isNaN(speed)) {
        // 构建T45格式的字符串
        const dataString = `S${speed}`;
        // 将字符串转换为ArrayBuffer
        const buffer = this.stringToArrayBuffer(dataString);
        
        wx.writeBLECharacteristicValue({
            deviceId: app.globalData.deviceId,
            serviceId: app.globalData.serviceId,
            characteristicId: app.globalData.characteristicId,
            value: buffer,
            success: (res) => {
                console.log('速度发送成功');
                wx.showToast({
                    title: '速度发送成功',
                    icon: 'success'
                });
            },
            fail: (err) => {
                console.error('速度发送失败:', err);
                wx.showToast({
                    title: '速度发送失败',
                    icon: 'none'
                });
            }
        });
    } else {
        wx.showToast({
            title: '请输入有效的阈值',
            icon: 'none'
        });
    }
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
    this.setspeed();
},
  bindInput: function (e) {
    this.setData({
      speed: e.detail.value
    });
  }
})

