const recorderManager = wx.getRecorderManager()
const options = {
    duration: 60000,//指定录音的时长，单位 ms
    sampleRate: 16000,//采样率
    numberOfChannels: 1,//录音通道数
    encodeBitRate: 48000,//编码码率
    format: 'pcm',//音频格式，有效值 aac/pcm
  }
  var filesize,tempFilePath
Page({
    /**
     * 页面的初始数据
     */
    data: {
        token:'',
        content:'',
        isRecording: false // 添加录音状态
    },

  onLoad(options) {
    //获取storge中的token
    let that=this;
    wx.getStorage({
        key:'expires_in',
        success(res){
            console.log("缓存中有access_token")
            console.log("token失效时间：",res.data)
            const newT = new Date().getTime();
            // 用当前时间和存储的时间判断，token是否已过期
            if (newT > parseInt(res.data)) {
                console.log("token过期，重新获取token")
                that.getToken();
            } else {
               that.getToken();
                console.log("获取本地缓存的token")
                that.setData({
                    token:wx.getStorageSync('access_token')
                });
            }
        },fail(){
            console.log("缓存中没有access_token")
            that.getToken();
        }
    });
  },

// 获取token
  getToken:function(){
    let that=this;
    let ApiKey='hILGfxaXbxQRhGI2pBQuFy61';//你自己的apikey
    let SecretKey='JnTVvZf2nfbHsNqYTnP8Z4D7OmkiVwXs';//你自己的SecretKey
    const url = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id='+ApiKey+'&client_secret='+SecretKey
    wx.request({
        url:url,
        method: 'POST',
        success(res){
          console.log("创建access_token成功",res)
            //将access_token存储到storage中
            wx.setStorage({
              key:'access_token',
              data:res.data.access_token
            });
            var date=new Date().getTime();
            let time=date+2592000*1000;
            console.log('三十天后的时间',time);
            wx.setStorage({
              key:'expires_in',
              data:time
            });
            that.setData({
                token:res.data.access_token
            });
            },
    });
  },
  //开始录音
  touchStart: function () {
    const that = this;
    wx.authorize({
      scope: 'scope.record',
      success() {
        console.log("录音授权成功");
        recorderManager.start(options);
        recorderManager.onStart(() => {
          console.log('recorder start');
          that.setData({ isRecording: true });
          wx.showToast({
            title: '开始录音',
            icon: 'success'
          });
        });
      },
      fail() {
        console.log("录音失败");
        wx.showToast({
          title: '录音授权失败',
          icon: 'none'
        });
      },
  })
},

  //停止录音
  touchEnd: function () {
    let that = this
    that.setData({ isRecording: false });
    recorderManager.stop();
    recorderManager.onStop((res) => {
    console.log('文件路径==', res)
    tempFilePath= res.tempFilePath;
    wx.showToast({
      title: '录音已停止',
      icon: 'success'
    });
    //获取文件长度
    wx.getFileSystemManager().getFileInfo({
        filePath: tempFilePath,
        success: function (res) {
          filesize = res.size
          console.log('文件长度', res)
          //that.shibie()
        }, fail: function (res) {   
          console.log("读取文件长度错误",res);
        }
      })   
    });
  },
 //语音识别
 shibie(){
  let that = this
  wx.getFileSystemManager().readFile({
    filePath: tempFilePath,
    encoding: 'base64',
    success: function (res) {
      wx.request({
        url: 'http://vop.baidu.com/server_api',
        data: {
          token: that.data.token,
          cuid: "12_56",
          format: 'pcm',
          rate: 16000,
          channel: 1,
          speech: res.data,
          len: filesize
        },
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        }, 
        method: "post",
        success: function (res) {
            if (res.data.result == '') {
                wx.showModal({
                  title: '提示',
                  content: '听不清楚，请重新说一遍！',
                  showCancel: false
                })
                return;
            }
            console.log("识别成功==",res.data);
            that.setData({
               content:res.data.result
            })
            if(res.data.result.toString().indexOf("蓝牙")>=0){
              
               this.bleInit();
            }
        },
        fail: function (res) {
          console.log("失败",res);
        }
      }); //语音识别结束
    }
  })     
 }
})