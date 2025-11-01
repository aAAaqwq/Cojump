// 云函数：语音识别
// 功能：将录音文件识别转为文字
const cloud = require('wx-server-sdk');
const tencentcloud = require('tencentcloud-sdk-nodejs');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });

const asrConfig = {
  httpProfile: {
    endpoint: "asr.tencentcloudapi.com",
  },
  region: '',
  credential: {
    secretId: process.env.TencentSecretId,
    secretKey: process.env.TencentSecretKey
  }
};

// 云函数入口函数
exports.main = async (event, context) => {
  const { fileID } = event; // 云存储文件ID

  try {
    //  从云存储获取文件下载URL
    const res = await cloud.getTempFileURL({
      fileList: [fileID],
    })
    const fileUrl = res.fileList[0].tempFileURL;
    console.log('文件下载URL:', fileUrl);
    //  获取腾讯云凭证（从环境变量读取）
    const secretId = process.env.TencentSecretId;
    const secretKey = process.env.TencentSecretKey;

    if (!secretId || !secretKey) {
      throw new Error('腾讯云凭证未配置，请在环境变量中设置 TencentSecretId 和 TencentSecretKey');
    }

    // 4. 创建腾讯云ASR客户端
    const AsrClient = tencentcloud.asr.v20190614.Client;
    const client = new AsrClient(asrConfig);

    // 5. 调用ASR API识别
    const params = {
      "EngSerViceType": "16k_zh-PY",
      "SourceType": 0,
      "VoiceFormat": "mp3",
      "Url": fileUrl,
    };

    console.log('调用腾讯云ASR API...');
    const response = await client.SentenceRecognition(params);
    console.log('识别响应:', response);

    // 6. 处理识别结果
    if (response.Result) {
      return {
        success: true,
        text: response.Result,
        message: '识别成功'
      };
    } else {
      return {
        success: false,
        text: '',
        message: '识别结果为空',
        error: response
      };
    }

  } catch (error) {
    console.error('语音识别失败:', error);
    return {
      success: false,
      text: '',
      message: error.message || '识别失败',
      error: error.toString()
    };
  }
};


