// pages/login/login.js
Page({

  /**
   * 页面的初始数据
   */
  data: {
    username: '',           // 用户名
    password: '',           // 密码
    showPassword: false,    // 是否显示密码
    rememberPassword: false, // 是否记住密码
    focusUsername: false,   // 用户名输入框是否聚焦
    focusPassword: false,   // 密码输入框是否聚焦
    loading: false,         // 登录加载状态
    platform: '',          // 当前平台
    isIOS: false,          // 是否为iOS平台
    isAndroid: false,      // 是否为安卓平台
    systemInfo: {}         // 系统信息
  },

  /**
   * 生命周期函数--监听页面加载
   */
  onLoad(options) {
    // 检测平台类型
    this.detectPlatform();
    // 尝试从本地存储读取记住的账号密码
    this.loadRememberedAccount();
  },

  /**
   * 检测平台类型
   */
  detectPlatform() {
    try {
      const systemInfo = wx.getSystemInfoSync();
      const platform = systemInfo.platform;
      const model = systemInfo.model || '';
      
      // 处理开发工具环境
      let isIOS = false;
      let isAndroid = false;
      let actualPlatform = platform;
      
      if (platform === 'devtools') {
        // 在开发工具中，通过model判断模拟的设备类型
        if (model.toLowerCase().includes('iphone') || model.toLowerCase().includes('ipad')) {
          isIOS = true;
          actualPlatform = 'ios';
          console.log('开发工具环境，检测到iOS模拟器');
        } else if (model.toLowerCase().includes('android') || model.toLowerCase().includes('samsung') || model.toLowerCase().includes('huawei') || model.toLowerCase().includes('xiaomi')) {
          isAndroid = true;
          actualPlatform = 'android';
          console.log('开发工具环境，检测到安卓模拟器');
        } else {
          // 默认按iOS处理（因为您提到要在iOS真机测试）
          isIOS = true;
          actualPlatform = 'ios';
          console.log('开发工具环境，默认按iOS处理');
        }
      } else {
        // 真机环境
        isIOS = platform === 'ios';
        isAndroid = platform === 'android';
      }
      
      console.log('平台检测结果:', {
        platform: platform,
        actualPlatform: actualPlatform,
        isIOS: isIOS,
        isAndroid: isAndroid,
        model: model,
        systemInfo: systemInfo
      });
      
      this.setData({
        platform: actualPlatform,
        isIOS: isIOS,
        isAndroid: isAndroid,
        systemInfo: systemInfo
      });
    } catch (error) {
      console.error('平台检测失败:', error);
      // 默认设置为iOS平台（因为您主要关心iOS真机测试）
      this.setData({
        platform: 'ios',
        isIOS: true,
        isAndroid: false
      });
    }
  },

  /**
   * 加载记住的账号密码
   */
  loadRememberedAccount() {
    try {
      const remembered = wx.getStorageSync('rememberedAccount');
      if (remembered) {
        this.setData({
          username: remembered.username || '',
          password: remembered.password || '',
          rememberPassword: true
        });
      }
    } catch (e) {
      console.error('读取记住的账号失败:', e);
    }
  },

  /**
   * 用户名输入事件
   */
  onUsernameInput(e) {
    this.setData({
      username: e.detail.value
    });
  },

  /**
   * 密码输入事件
   */
  onPasswordInput(e) {
    this.setData({
      password: e.detail.value
    });
  },

  /**
   * 用户名输入框聚焦
   */
  onUsernameFocus() {
    this.setData({
      focusUsername: true
    });
  },

  /**
   * 用户名输入框失焦
   */
  onUsernameBlur() {
    this.setData({
      focusUsername: false
    });
  },

  /**
   * 密码输入框聚焦
   */
  onPasswordFocus() {
    this.setData({
      focusPassword: true
    });
  },

  /**
   * 密码输入框失焦
   */
  onPasswordBlur() {
    this.setData({
      focusPassword: false
    });
  },

  /**
   * 切换密码显示/隐藏 - 平台兼容方案
   */
  togglePassword(e) {
    const newShowPassword = !this.data.showPassword;
    
    // console.log('=== 密码切换调试信息 ===');
    // console.log('当前状态:', this.data.showPassword);
    // console.log('新状态:', newShowPassword);
    // console.log('密码内容:', this.data.password);
    // console.log('密码长度:', this.data.password.length);
    // console.log('当前平台:', this.data.platform);
    // console.log('是否iOS:', this.data.isIOS);
    // console.log('是否安卓:', this.data.isAndroid);
    
    this.setData({
      showPassword: newShowPassword
    });
    
    // 平台特定的处理
    this.handlePlatformSpecificToggle(newShowPassword);
    
    // 添加视觉反馈
    this.showPasswordToggleFeedback(newShowPassword);
    
    // 强制刷新页面状态
    wx.nextTick(() => {
      console.log('=== 状态更新后 ===');
      console.log('showPassword:', this.data.showPassword);
      // console.log('password:', this.data.password);
    });
  },

  /**
   * 平台特定的切换处理
   */
  handlePlatformSpecificToggle(showPassword) {
    if (this.data.isIOS) {
      // iOS平台：使用-webkit-text-security
      console.log('iOS平台密码切换处理');
      this.handleIOSPasswordToggle(showPassword);
    } else if (this.data.isAndroid) {
      // 安卓平台：使用双input切换
      console.log('安卓平台密码切换处理');
      this.handleAndroidPasswordToggle(showPassword);
    } else {
      // 其他平台：使用标准方案
      console.log('其他平台密码切换处理');
      this.handleDefaultPasswordToggle(showPassword);
    }
  },

  /**
   * iOS平台密码切换处理
   */
  handleIOSPasswordToggle(showPassword) {
    // iOS平台使用-webkit-text-security，无需额外处理
    console.log('iOS密码切换完成，使用-webkit-text-security');
  },

  /**
   * 安卓平台密码切换处理
   */
  handleAndroidPasswordToggle(showPassword) {
    // 安卓平台使用双input切换，无需额外处理
    console.log('安卓密码切换完成，使用双input切换');
  },

  /**
   * 默认平台密码切换处理
   */
  handleDefaultPasswordToggle(showPassword) {
    // 默认平台使用标准方案，无需额外处理
    console.log('默认平台密码切换完成，使用标准方案');
  },

  /**
   * 开发工具测试：iOS模式
   */
  testIOSMode() {
    console.log('切换到iOS测试模式');
    this.setData({
      platform: 'ios',
      isIOS: true,
      isAndroid: false
    });
    wx.showToast({
      title: '已切换到iOS模式',
      icon: 'none',
      duration: 1500
    });
  },

  /**
   * 开发工具测试：安卓模式
   */
  testAndroidMode() {
    console.log('切换到安卓测试模式');
    this.setData({
      platform: 'android',
      isIOS: false,
      isAndroid: true
    });
    wx.showToast({
      title: '已切换到安卓模式',
      icon: 'none',
      duration: 1500
    });
  },

  /**
   * 开发工具测试：重置平台检测
   */
  resetPlatform() {
    console.log('重置平台检测');
    this.detectPlatform();
    wx.showToast({
      title: '已重置平台检测',
      icon: 'none',
      duration: 1500
    });
  },

  /**
   * 显示密码切换反馈
   */
  showPasswordToggleFeedback(showPassword) {
    const message = showPassword ? '密码已显示' : '密码已隐藏';
    
    // 使用更轻量的反馈方式
    wx.showToast({
      title: message,
      icon: 'none',
      duration: 800
    });
  },

  /**
   * 切换记住密码
   */
  toggleRemember() {
    this.setData({
      rememberPassword: !this.data.rememberPassword
    });
  },

  /**
   * 忘记密码
   */
  onForgotPassword() {
    wx.showToast({
      title: '请联系管理员重置密码',
      icon: 'none',
      duration: 2000
    });
  },

  /**
   * 登录验证
   */
  validateLogin() {
    const { username, password } = this.data;

    if (!username) {
      wx.showToast({
        title: '请输入账号',
        icon: 'none',
        duration: 2000
      });
      return false;
    }

    if (!password) {
      wx.showToast({
        title: '请输入密码',
        icon: 'none',
        duration: 2000
      });
      return false;
    }

    if (username.length < 3) {
      wx.showToast({
        title: '账号长度不能少于3位',
        icon: 'none',
        duration: 2000
      });
      return false;
    }

    if (password.length < 6) {
      wx.showToast({
        title: '密码长度不能少于6位',
        icon: 'none',
        duration: 2000
      });
      return false;
    }

    return true;
  },

  /**
   * 保存记住的账号密码
   */
  saveRememberedAccount() {
    if (this.data.rememberPassword) {
      try {
        wx.setStorageSync('rememberedAccount', {
          username: this.data.username,
          password: this.data.password
        });
      } catch (e) {
        console.error('保存记住的账号失败:', e);
      }
    } else {
      try {
        wx.removeStorageSync('rememberedAccount');
      } catch (e) {
        console.error('删除记住的账号失败:', e);
      }
    }
  },

  /**
   * 登录按钮点击事件
   */
  async onLogin() {
    // 验证输入
    if (!this.validateLogin()) {
      return;
    }

    // 设置加载状态
    this.setData({
      loading: true
    });

    try {
      // 这里应该调用后端登录接口
      // 示例：使用云函数进行登录验证
      const { username, password } = this.data;

      // 模拟登录请求（实际项目中应替换为真实的登录逻辑）
      const result = await wx.cloud.callFunction({
        name: 'login',
        data: {
          username,
          password
        }
      });

      // 模拟登录延迟
      // await new Promise(resolve => setTimeout(resolve, 1500));

      // 模拟登录成功（实际项目中应根据后端返回判断）
      // const mockSuccess = true;

      console.log("登录结果:",result)
      if (result.result.success) {
        // 保存登录状态
        wx.setStorageSync('isLogin', true);
        wx.setStorageSync('userInfo', {
          username: username,
          loginTime: Date.now()
        });

        // 保存记住的账号
        this.saveRememberedAccount();

        // 显示成功提示
        wx.showToast({
          title: '登录成功',
          icon: 'success',
          duration: 1500
        });

        // 跳转到首页
        setTimeout(() => {
          wx.reLaunch({
            url: '/pages/home/home'
          });
        }, 1500);
      } else {
        throw new Error('账号或密码错误');
      }

    } catch (error) {
      console.error('登录失败:', error);
      wx.showToast({
        title: error.message || '登录失败，请重试',
        icon: 'none',
        duration: 2000
      });
    } finally {
      // 取消加载状态
      this.setData({
        loading: false
      });
    }
  }

})