/**
 * 小程序更新管理
 * 检测小程序是否有新版本，自动更新
 */
function updateManager() {
  // 判断微信基础库版本是否支持 UpdateManager
  if (!wx.canIUse('getUpdateManager')) {
    console.log('当前微信版本过低，无法使用更新功能');
    return;
  }

  const updateManager = wx.getUpdateManager();

  // 检查更新
  updateManager.onCheckForUpdate((res) => {
    if (res.hasUpdate) {
      console.log('发现新版本');
    }
  });

  // 新版本下载成功
  updateManager.onUpdateReady(() => {
    wx.showModal({
      title: '更新提示',
      content: '新版本已经准备好，是否重启应用？',
      success(res) {
        if (res.confirm) {
          // 新版本下载完成，调用 applyUpdate 应用新版本并重启
          updateManager.applyUpdate();
        }
      },
    });
  });

  // 新版本下载失败
  updateManager.onUpdateFailed(() => {
    wx.showModal({
      title: '更新失败',
      content: '新版本下载失败，请检查网络后重试',
      showCancel: false,
    });
  });
}

export default updateManager;

