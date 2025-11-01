// pages/VoiceRecognition/VoiceRecognition.js
const recorderManager = wx.getRecorderManager()
const app = getApp();
const options = {
  duration: 15000,//指定录音的时长，单位 ms
  sampleRate: 16000,//采样率
  numberOfChannels: 1,//录音通道数
  encodeBitRate: 48000,//编码码率
  format: 'mp3',//音频格式，有效值 mp3/wav/aac/pcm
  // frameSize: 50,//指定帧大小，单位 ms
}

Page({
  /**
   * 页面的初始数据
   */
  data: {
    tempFilePath: '',
    filesize: 0,
    isRecording: false,
  },

  /**
   * 页面的事件处理函数
   */

  // 开始录音
  StartRecording: function () {
    let that = this;
    wx.authorize({
      scope: 'scope.record',
      success() {
        console.log("录音授权成功");
        recorderManager.start(options);
        recorderManager.onStart(() => {
          wx.showToast({
            title: '开始录音',
            icon: 'success'
          });
          that.setData({
            isRecording: true,
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
    });
  },

  // 停止录音
  StopRecording: function () {
    let that = this;
    recorderManager.stop();
    recorderManager.onStop((res) => {
      console.log('停止录音:', res);
      that.setData({
        tempFilePath: res.tempFilePath,
        isRecording: false,
      });
      wx.showToast({
        title: '录音已停止',
        icon: 'success'
      });
      //获取文件长度
      wx.getFileSystemManager().getFileInfo({
        filePath: that.data.tempFilePath,
        success: function (res) {
          filesize = res.size;
          console.log('文件长度', res);
        },
      });

      // 上传文件
      this.UploadFile(this.data.tempFilePath);
    });
  },

  // 上传文件
  UploadFile: function (tempFilePath) {
    // 上传到云存储
    wx.cloud.uploadFile({
      cloudPath: 'voice/' +app.globalData.openId + '/' + Date.now() + '.mp3',
      filePath: tempFilePath,
      success: function (res) {
        console.log('上传成功', res);
        // 上传成功后，调用云函数进行语音识别        
        this.VoiceRecognize(res.fileID);
      },
      fail: function (err) {
        console.log('上传失败', err);
      },
    });
  },

  // 语音识别
  VoiceRecognize: function (fileId) {
    try {
      const res = wx.cloud.callFunction({
        name: 'VoiceRecognition',
        data: {
          fileId: fileId,
        },
      });
      console.log('语音识别成功', res);
      return res.result.data;
    } catch (error) {
      console.log('语音识别失败', error);
    }
  },


})