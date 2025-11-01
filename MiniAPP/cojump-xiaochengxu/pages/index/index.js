// index.js
// 获取应用实例
const app = getApp()
const recorderManager = wx.getRecorderManager()
const options = {
    duration: 60000,//指定录音的时长，单位 ms
    sampleRate: 16000,//采样率
    numberOfChannels: 1,//录音通道数
    encodeBitRate: 48000,//编码码率
    format: 'pcm',//音频格式，有效值 aac/pcm
  }
  var filesize,tempFilePath

  let isConnected = false;

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
    token:'',
    content:'',
    userInfo: null, // 用户信息
    isConnected: false,
  },

  onLoad() {
    // this.chartComponent = this.selectComponent('#chartCanvas');
    // 加载用户信息
    this.loadUserInfo();
  },

  onShow() {
    // 页面显示时重新加载用户信息
    this.loadUserInfo();
    this.setData({ isConnected });
  },

  // 加载用户信息
  loadUserInfo() {
    const userInfo = wx.getStorageSync('userInfo');
    if (userInfo) {
      this.setData({ userInfo });
    }
  },

  // 跳转到个人信息页面
  toUserInfo() {
    wx.navigateTo({
      url: '/pages/userInfo/userInfo'
    });
  },
  
  BLE: function() {
    if (isConnected){return;}
    console.log("BLE函数被调用");
    this.bleInit();
  },//蓝牙连接

  bleInit() {
    wx.showToast({
      title: '搜索蓝牙设备',
      icon: 'loading',
      duration: 20000 // 显示时间，单位ms
    });
    console.log('searchBle')
    // 监听扫描到新设备事件
    
    wx.onBluetoothDeviceFound((res) => {
      res.devices.forEach((device) => {
        // 这里可以做一些过滤
        console.log('Device Found', device)
        if(device.name == "HAND"){
          // 找到设备开始连接
          this.bleConnection(device.deviceId);
          wx.stopBluetoothDevicesDiscovery()
        }
      })
      // 找到要搜索的设备后，及时停止扫描
      // 
    })

    // 初始化蓝牙模块
    wx.openBluetoothAdapter({
      mode: 'central',
      success: (res) => {
        // 开始搜索附近的蓝牙外围设备
        wx.startBluetoothDevicesDiscovery({
          allowDuplicatesKey: true,
        })
      },
      fail: (res) => {
        if (res.errCode !== 10001) return
        wx.onBluetoothAdapterStateChange((res) => {
          if (!res.available) return
          // 开始搜寻附近的蓝牙外围设备
          wx.startBluetoothDevicesDiscovery({
            allowDuplicatesKey: false,
          })
        })
      }
    })
    var that = this
    wx.onBLECharacteristicValueChange((result) => {
      console.log('onBLECharacteristicValueChange',result.value)
      let hex = that.ab2hex(result.value)
      console.log('hextoString',that.hextoString(hex))
      console.log('hex',hex)
    
      const dataStr = that.hextoString(hex)
      console.log('Received data:', dataStr)
      
      const app = getApp();
      app.globalData.eventBus.$emit('bleDataUpdate', dataStr);
      
    })
  },
  bleConnection(deviceId){
    wx.createBLEConnection({
      deviceId, // 搜索到设备的 deviceId
      success: () => {
        // 连接成功，获取服务
        console.log('连接成功，获取服务')
        this.bleGetDeviceServices(deviceId)
      }
    })
    var that = this;
    wx.onBLEConnectionStateChange(function(res) {
      that.setData({isConnected:!!res.connected});
      if (res.connected) {
        // 连接成功
        isConnected = true;
        wx.showToast({
          title: '蓝牙设备连接成功',
          icon: 'success',
          duration: 2000 // 显示时间，单位ms
        });
      } else {
        // 连接断开
        isConnected = false;
        console.log('蓝牙连接断开');
        // 这里可以添加处理断开连接的逻辑
        // 比如显示一个Toast提示用户蓝牙已断开
        wx.showToast({
          title: '蓝牙断开，请退回主界面重新连接',
          icon: 'none',
          duration: 4000
        });
        // 尝试重新连接或其他操作...
      }
    });
  },
  bleGetDeviceServices(deviceId){
    wx.getBLEDeviceServices({
      deviceId, // 搜索到设备的 deviceId
      success: (res) => {
        console.log(res.services)
        for (let i = 0; i < res.services.length; i++) {
          if (res.services[i].isPrimary) {
            // 可根据具体业务需要，选择一个主服务进行通信
            this.bleGetDeviceCharacteristics(deviceId,res.services[i].uuid)
            break;
          }
        }
      }
    })
  },
  bleGetDeviceCharacteristics(deviceId,serviceId){
    wx.getBLEDeviceCharacteristics({
      deviceId, // 搜索到设备的 deviceId
      serviceId, // 上一步中找到的某个服务
      success: (res) => {
        for (let i = 0; i < res.characteristics.length; i++) {
          let item = res.characteristics[i]
          console.log(item)
          if (item.properties.write) { // 该特征值可写
            // 本示例是向蓝牙设备发送一个 0x00 的 16 进制数据
            // 实际使用时，应根据具体设备协议发送数据
            // let buffer = new ArrayBuffer(1)
            // let dataView = new DataView(buffer)
            // dataView.setUint8(0, 0)
            // let senddata = 'FF';
            // let buffer = this.hexString2ArrayBuffer(senddata);
            var buffer = this.stringToBytes("getid")
            this.setData({
              'deviceId':deviceId,
              'serviceId':serviceId,
              'characteristicId':item.uuid
            })
            app.globalData.deviceId = deviceId; // 更新全局变量 deviceId
            app.globalData.serviceId = serviceId; // 更新全局变量 serviceId
            app.globalData.characteristicId = item.uuid;
            console.log(app.globalData.deviceId);
            console.log(app.globalData.serviceId);
            console.log(app.globalData.characteristicId);
            wx.writeBLECharacteristicValue({
              deviceId,
              serviceId,
              characteristicId: item.uuid,
              value: buffer,
            })
          }
          if (item.properties.read) { // 改特征值可读
            wx.readBLECharacteristicValue({
              deviceId,
              serviceId,
              characteristicId: item.uuid,
            })
          }
          if (item.properties.notify || item.properties.indicate) {
            // 必须先启用 wx.notifyBLECharacteristicValueChange 才能监听到设备 onBLECharacteristicValueChange 事件
            wx.notifyBLECharacteristicValueChange({
              deviceId,
              serviceId,
              characteristicId: item.uuid,
              state: true,
            })
          }
        }
      }
    })
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
  
  ToPage1: function() {
    wx.navigateTo({
      url: '/pages/page1/page1'
    });
  console.log('跳转至肌力训练模式')
  },

  ToPage2: function() {
    wx.navigateTo({
      url: '/pages/page2/page2'
    });
    var buffer = this.stringToBytes("EMG")
    wx.writeBLECharacteristicValue({
      deviceId:app.globalData.deviceId,
      serviceId:app.globalData.serviceId,
      characteristicId:app.globalData.characteristicId,
      value: buffer,
    })
  console.log('跳转至EMG模式')  
  },

  ToVoiceRecognition: function() {
    wx.navigateTo({
      url: '/pages/VoiceRecognition/VoiceRecognition'
    });
  console.log('语音识别')
  },

  ToPage3: function() {
    wx.navigateTo({
      url: '/pages/RehabilitationHistory/RehabilitationHistory'
    });
  console.log('个人康复历史记录')
  },

  ToPage4: function() {
    wx.navigateTo({
      url: '/pages/EMGxuanze/EMGxuanze'
    });
  console.log('EMG训练')
  },
  


})