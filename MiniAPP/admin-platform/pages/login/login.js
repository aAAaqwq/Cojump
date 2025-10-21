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
    loading: false          // 登录加载状态
  },

  /**
   * 生命周期函数--监听页面加载
   */
  onLoad(options) {
    // 尝试从本地存储读取记住的账号密码
    this.loadRememberedAccount();
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
   * 切换密码显示/隐藏
   */
  togglePassword() {
    this.setData({
      showPassword: !this.data.showPassword
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