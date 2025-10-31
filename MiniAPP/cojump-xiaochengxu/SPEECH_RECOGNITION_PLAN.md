# 语音识别功能技术规划文档

## 📋 功能概述

实现语音识别功能，识别用户的语音命令并跳转到相应功能页面。

**当前阶段目标**：先实现语音识别，将语音转换为文字。

---

## 🎯 技术方案对比

### 方案一：微信小程序录音 + 腾讯云ASR（推荐⭐）

**优势：**
- ✅ 与现有云开发环境无缝集成
- ✅ 免费额度充足：每月10000分钟
- ✅ 识别准确率高，支持多种方言
- ✅ 响应速度快，延迟低
- ✅ 无需额外申请第三方账号

**成本：**
- 免费额度：10000分钟/月
- 超出后：0.036元/分钟

**技术栈：**
- 前端：`wx.startRecord()` + `wx.stopRecord()`
- 云函数：调用腾讯云ASR API

---

### 方案二：微信小程序录音 + 百度语音识别

**优势：**
- ✅ 免费额度：每天5万次调用
- ✅ API简单易用
- ✅ 准确率较高

**劣势：**
- ❌ 需要申请百度账号和API Key
- ❌ 跨平台调用，可能有一定延迟

**成本：**
- 免费：每天50000次
- 超出后：0.0048元/次

---

### 方案三：微信小程序录音 + 讯飞语音

**优势：**
- ✅ 专业语音识别服务
- ✅ 支持多种语言和方言

**劣势：**
- ❌ 免费额度较少
- ❌ 需要企业认证
- ❌ 接入复杂度较高

---

### 方案四：使用微信小程序插件

**优势：**
- ✅ 快速接入
- ✅ 可能有现成UI

**劣势：**
- ❌ 可能收费
- ❌ 定制化程度低
- ❌ 依赖第三方维护

---

## ✅ 推荐方案：方案一（腾讯云ASR）

基于项目已使用微信云开发，**推荐使用腾讯云ASR方案**。

---

## 📦 技术架构

```
┌─────────────────┐
│   小程序前端     │
│  ┌───────────┐  │
│  │ 录音按钮   │  │  wx.startRecord()
│  └───────────┘  │  wx.stopRecord()
└────────┬────────┘
         │ 上传录音文件
         ▼
┌─────────────────┐
│   云存储         │  临时存储录音文件
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   云函数         │  speechRecognition
│   ┌───────────┐ │
│   │ 下载录音   │ │
│   │ 调用ASR   │ │  腾讯云ASR API
│   │ 返回文字   │ │
│   └───────────┘ │
└────────┬────────┘
         │ 返回识别结果
         ▼
┌─────────────────┐
│   小程序前端     │
│  ┌───────────┐  │
│  │ 显示结果   │  │  展示识别文字
│  │ 匹配命令   │  │  匹配关键词
│  │ 页面跳转   │  │  执行跳转逻辑
│  └───────────┘  │
└─────────────────┘
```

---

## 🛠️ 实现步骤

### 第一阶段：基础语音识别（当前目标）

#### 1. 前端录音功能

**需要权限：**
```json
// app.json
{
  "permission": {
    "scope.record": {
      "desc": "需要使用录音功能进行语音识别"
    }
  }
}
```

**核心API：**
- `wx.authorize({ scope: 'scope.record' })` - 申请录音权限
- `wx.startRecord()` - 开始录音
- `wx.stopRecord()` - 停止录音，返回临时文件路径

#### 2. 云函数：语音识别

**创建云函数：** `cloudfunctions/speechRecognition/`

**依赖包：**
```json
{
  "dependencies": {
    "wx-server-sdk": "~2.6.3",
    "@tencentcloud-sdk/capi": "^3.0.0"
  }
}
```

**核心功能：**
1. 接收前端上传的录音文件路径
2. 从云存储下载录音文件
3. 调用腾讯云ASR API进行识别
4. 返回识别结果文本

#### 3. 前端语音识别组件

创建可复用的语音识别组件或工具函数。

---

### 第二阶段：命令匹配与跳转

#### 1. 关键词匹配规则

```javascript
const commandMap = {
  '模式1': '/pages/page1/page1',
  '模式2': '/pages/page2/page2',
  '肌电测量': '/pages/EMGceliang/EMGceliang',
  '主动训练': '/pages/EMGtrain/EMGtrain',
  '阈值测量': '/pages/EMGceliang/EMGceliang',
  '历史记录': '/pages/RehabilitationHistory/RehabilitationHistory',
  // ... 更多命令
}
```

#### 2. 智能匹配算法

- 简单匹配：精确匹配关键词
- 模糊匹配：使用编辑距离算法
- 同义词识别：扩展关键词库

---

## 📁 文件结构规划

```
cojump-xiaochengxu/
├── pages/
│   └── voiceRecognition/          # 语音识别页面（可选）
│       ├── voiceRecognition.js
│       ├── voiceRecognition.wxml
│       └── voiceRecognition.wxss
├── components/                     # 语音识别组件（推荐）
│   └── voice-input/
│       ├── voice-input.js
│       ├── voice-input.wxml
│       └── voice-input.wxss
├── utils/
│   └── voiceHelper.js             # 语音识别工具函数
├── cloudfunctions/
│   └── speechRecognition/         # 语音识别云函数
│       ├── index.js
│       ├── package.json
│       └── config.json
└── app.json                        # 添加录音权限配置
```

---

## 🔧 核心技术实现

### 1. 前端录音实现

```javascript
// utils/voiceHelper.js
class VoiceRecorder {
  constructor() {
    this.isRecording = false;
    this.recordManager = wx.getRecorderManager();
    this.initRecorder();
  }

  initRecorder() {
    this.recordManager.onStart(() => {
      this.isRecording = true;
      console.log('开始录音');
    });

    this.recordManager.onStop((res) => {
      this.isRecording = false;
      const { tempFilePath, duration } = res;
      console.log('录音结束', tempFilePath, duration);
      this.uploadAndRecognize(tempFilePath);
    });

    this.recordManager.onError((err) => {
      console.error('录音错误', err);
      wx.showToast({
        title: '录音失败',
        icon: 'none'
      });
    });
  }

  startRecord() {
    // 检查权限
    wx.authorize({
      scope: 'scope.record',
      success: () => {
        this.recordManager.start({
          duration: 60000, // 最长60秒
          sampleRate: 16000,
          numberOfChannels: 1,
          encodeBitRate: 96000,
          format: 'mp3'
        });
      },
      fail: () => {
        wx.showModal({
          title: '需要录音权限',
          content: '请在设置中开启录音权限',
          showCancel: false
        });
      }
    });
  }

  stopRecord() {
    if (this.isRecording) {
      this.recordManager.stop();
    }
  }

  async uploadAndRecognize(filePath) {
    wx.showLoading({ title: '识别中...' });
    
    try {
      // 上传到云存储
      const cloudPath = `voice/${Date.now()}.mp3`;
      const uploadRes = await wx.cloud.uploadFile({
        cloudPath,
        filePath
      });

      // 调用云函数识别
      const res = await wx.cloud.callFunction({
        name: 'speechRecognition',
        data: {
          fileID: uploadRes.fileID
        }
      });

      wx.hideLoading();
      
      if (res.result.success) {
        const text = res.result.text;
        console.log('识别结果:', text);
        return text;
      } else {
        throw new Error(res.result.error);
      }
    } catch (error) {
      wx.hideLoading();
      wx.showToast({
        title: '识别失败',
        icon: 'none'
      });
      console.error('识别错误', error);
    }
  }
}

module.exports = VoiceRecorder;
```

---

### 2. 云函数：腾讯云ASR识别

```javascript
// cloudfunctions/speechRecognition/index.js
const cloud = require('wx-server-sdk');
const tencentcloud = require('tencentcloud-sdk-nodejs');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });

const AsrClient = tencentcloud.asr.v20190614.Client;

exports.main = async (event, context) => {
  const { fileID } = event;

  try {
    // 1. 下载录音文件
    const file = await cloud.downloadFile({ fileID });
    const buffer = file.fileContent;

    // 2. 转换为base64
    const base64Audio = buffer.toString('base64');

    // 3. 调用腾讯云ASR API
    const client = new AsrClient({
      credential: {
        secretId: process.env.TENCENT_SECRET_ID, // 从环境变量读取
        secretKey: process.env.TENCENT_SECRET_KEY
      },
      region: 'ap-beijing'
    });

    const params = {
      ProjectId: 0,
      SubServiceType: 2, // 实时语音识别
      EngSerViceType: '16k_zh', // 中文16k采样率
      SourceType: 0, // 语音输入源：0-音频上传
      VoiceFormat: 'mp3',
      Data: base64Audio,
      DataLen: buffer.length
    };

    const response = await client.SentenceRecognition(params);
    
    if (response.Result) {
      return {
        success: true,
        text: response.Result
      };
    } else {
      return {
        success: false,
        error: '识别失败'
      };
    }
  } catch (error) {
    console.error('ASR识别错误', error);
    return {
      success: false,
      error: error.message
    };
  }
};
```

---

### 3. 命令匹配与跳转

```javascript
// utils/commandMatcher.js
class CommandMatcher {
  constructor() {
    this.commandMap = {
      '模式1': '/pages/page1/page1',
      '模式2': '/pages/page2/page2',
      '肌电测量': '/pages/EMGceliang/EMGceliang',
      '主动训练': '/pages/EMGtrain/EMGtrain',
      '阈值测量': '/pages/EMGceliang/EMGceliang',
      '功能选择': '/pages/EMGxuanze/EMGxuanze',
      '历史记录': '/pages/RehabilitationHistory/RehabilitationHistory',
      '首页': '/pages/index/index',
      '个人中心': '/pages/userInfo/userInfo'
    };

    // 同义词扩展
    this.synonyms = {
      '模式1': ['模式一', '第一种模式', '训练模式一'],
      '肌电测量': ['EMG测量', '肌电', '测量'],
      '主动训练': ['训练', '开始训练']
    };
  }

  match(text) {
    // 1. 精确匹配
    if (this.commandMap[text]) {
      return this.commandMap[text];
    }

    // 2. 模糊匹配
    for (const [key, path] of Object.entries(this.commandMap)) {
      if (text.includes(key)) {
        return path;
      }
    }

    // 3. 同义词匹配
    for (const [key, synonyms] of Object.entries(this.synonyms)) {
      if (synonyms.some(syn => text.includes(syn))) {
        return this.commandMap[key];
      }
    }

    return null;
  }

  navigate(text) {
    const path = this.match(text);
    if (path) {
      wx.navigateTo({ url: path });
      wx.showToast({
        title: '已跳转',
        icon: 'success'
      });
      return true;
    } else {
      wx.showToast({
        title: '未识别到有效命令',
        icon: 'none'
      });
      return false;
    }
  }
}

module.exports = CommandMatcher;
```

---

## 📝 使用示例

### 在页面中使用语音识别

```javascript
// pages/index/index.js
const VoiceRecorder = require('../../utils/voiceHelper');
const CommandMatcher = require('../../utils/commandMatcher');

Page({
  data: {
    isRecording: false,
    recognizedText: ''
  },

  onLoad() {
    this.voiceRecorder = new VoiceRecorder();
    this.commandMatcher = new CommandMatcher();
  },

  // 开始录音
  startVoice() {
    this.setData({ isRecording: true });
    this.voiceRecorder.startRecord();
  },

  // 停止录音并识别
  stopVoice() {
    this.setData({ isRecording: false });
    this.voiceRecorder.stopRecord();
  },

  // 处理识别结果
  async handleRecognition(text) {
    this.setData({ recognizedText: text });
    
    // 匹配命令并跳转
    this.commandMatcher.navigate(text);
  }
});
```

---

## 🔐 环境配置

### 1. 腾讯云密钥配置

在云函数环境变量中配置：
- `TENCENT_SECRET_ID` - 腾讯云SecretId
- `TENCENT_SECRET_KEY` - 腾讯云SecretKey

### 2. 申请腾讯云ASR服务

1. 登录 [腾讯云控制台](https://console.cloud.tencent.com/)
2. 开通「语音识别」服务
3. 创建API密钥（SecretId和SecretKey）
4. 在云函数环境变量中配置密钥

---

## 💰 成本估算

### 腾讯云ASR成本
- **免费额度**：10000分钟/月
- **超出后**：0.036元/分钟
- **示例**：假设每天识别100次，每次5秒，每月约250分钟，**完全免费**

---

## 📚 技术文档参考

1. [微信小程序录音API](https://developers.weixin.qq.com/miniprogram/dev/api/media/recorder/wx.startRecord.html)
2. [腾讯云语音识别文档](https://cloud.tencent.com/document/product/1093)
3. [微信云开发文档](https://developers.weixin.qq.com/miniprogram/dev/wxcloud/basis/getting-started.html)

---

## ✅ 实施检查清单

- [ ] 在`app.json`中添加录音权限
- [ ] 创建`utils/voiceHelper.js`录音工具类
- [ ] 创建`utils/commandMatcher.js`命令匹配类
- [ ] 创建`cloudfunctions/speechRecognition/`云函数
- [ ] 配置腾讯云ASR密钥到云函数环境变量
- [ ] 在目标页面集成语音识别功能
- [ ] 测试录音和识别功能
- [ ] 测试命令匹配和页面跳转
- [ ] 优化识别准确率和用户体验

---

## 🎯 后续优化方向

1. **离线识别**：使用本地语音识别SDK（需额外开发）
2. **实时识别**：使用流式识别，边录边识别
3. **多语言支持**：识别英文命令
4. **语音唤醒**：实现"小助手"唤醒词
5. **语音反馈**：识别成功后语音提示
6. **错误处理**：网络异常、识别失败等场景优化

---

## 📞 技术支持

如遇问题，参考：
- 腾讯云ASR技术支持：https://cloud.tencent.com/document/product/1093
- 微信小程序社区：https://developers.weixin.qq.com/community

---

**文档版本**：v1.0  
**最后更新**：2025-01-XX  
**维护者**：技术团队

