import updateManager from './common/updateManager';
import { init } from '@cloudbase/wx-cloud-client-sdk';

wx.cloud.init({
  env: 'cloudbase-7gky5kj9737551d9', // 指定云开发环境 ID
});

const client = init(wx.cloud);
const models = client.models;
globalThis.dataModel = models;
// 接下来就可以调用 models 上的数据模型增删改查等方法了

App({
  globalData: {
    isLogin: false,
    userInfo: null,
    loginTime: null,
  },
  onLaunch: function () {},
  onShow: function () {
    updateManager();
  },
});
