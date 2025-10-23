const app = getApp();

Page({
  // 页面配置 - 禁用滚动避免Canvas定位问题
  config: {
    disableScroll: true,
  },
  
  onShareAppMessage: function (res) {
    return {
      title: 'EMG阈值历史趋势图',
      path: '/pages/tuxiang/tuxiang',
      success: function () { },
      fail: function () { }
    }
  },
  data: {
    loading: false,
    hasData: false,
    emgData: {}, // 改为对象，key为日期字符串，value为EMG阈值
    // 日历相关
    currentYear: new Date().getFullYear(), // 当前年份
    currentMonth: new Date().getMonth() + 1, // 当前月份
    calendarDays: [], // 日历数据
    selectedDate: null, // 选中的日期（存储dateKey）
    selectedDateInfo: null, // 选中日期的详细信息:dateKey,date,year,month,day,weekDay,emgValue,hasEMGData
    // Canvas弹窗相关
    showChartModal: false,
    chartModalData: {
      title: '',
      dates: [],
      values: [],
      type: 'line',
      average: 0,
      max: 0,
      min: 0
    },
    // 图表交互状态
    isScaling: false,
    isPanning: false,
    chartScale: 1,
    chartPanX: 0,
    chartPanY: 0,
    initialDistance: 0,
    initialScale: 1,
    lastTouchX: 0,
    lastTouchY: 0,
    initialPanX: 0,
    initialPanY: 0,
    // 弹窗状态
    modalInitialized: false,
    // 患者友好视图
    showPatientView: true, // 新增：患者友好视图
    currentStatus: {
      todayValue: 0,
      yesterdayValue: 0,
      improvement: 0,
      status: 'good', // good, stable, attention
      message: '继续努力！'
    },
    weeklyProgress: {
      average: 0,
      target: 50,
      completion: 0
    },
    achievements: [],
    // 统计相关
    totalTrainingDays: 0,
    bestValue: 0
  },


  onLoad() {
    this.loadEMGData();
  },

  onShow() {
    // 页面显示时的逻辑
  },

  onReady() {
    // 页面准备完成时的逻辑
  },


  // 加载EMG历史数据
  loadEMGData() {
    this.setData({ loading: true });
    
    // 模拟数据生成（实际项目中替换为云函数调用）
    // setTimeout(() => {
    //   const sampleData = this.generateSampleEMGData();
    //   this.processEMGData(sampleData);
    //   this.setData({ hasData: true, loading: false });
      
    // }, 1000);

    // 实际的云函数调用（暂时注释）
    ///*
    wx.cloud.callFunction({
      name: 'getEMG',
      data: {
        openId: app.globalData.openId
      },
      success: (res) => {
        console.log('EMG数据获取成功:', res);
        if (res.result.success && res.result.data.length > 0) {
          // 处理获取的EMG数据
          const thresholdData = this.extractEMGData(res.result.data);
          this.processEMGData(thresholdData);
          this.setData({ hasData: true });
        } else {
          this.setData({ hasData: false });
          wx.showToast({
            title: '暂无EMG历史数据',
            icon: 'none',
            duration: 2000
          });
        }
        this.setData({ loading: false });
      },
      fail: (err) => {
        console.error('EMG数据获取失败:', err);
        this.setData({ loading: false, hasData: false });
        wx.showToast({
          title: '数据加载失败',
          icon: 'none'
        });
      }
    });
    //*/
  },
  // 提取EMG阈值数据
  extractEMGData(data) {
    const thresholdData = {};
    data.forEach(item => {
      // 将recordTime转换为日期
      const date = new Date(item.recordTime);
      const dateKey = this.formatDateKey(date);
      
      // 确保 emgThreshold 转换为数字类型
      const numericValue = parseFloat(item.emgThreshold);
      if (!isNaN(numericValue)) {
        thresholdData[dateKey] = numericValue;
      } else {
        console.warn('无效的EMG阈值数据:', item.emgThreshold, '日期:', dateKey);
      }
    });
    return thresholdData;
  },

   // 处理EMG阈值数据，转换为日历格式
   processEMGData(rawData) {
    console.log("处理EMG阈值数据:",rawData)
    // 保存所有原始数据
    this.setData({ emgData: rawData });
    this.setData({ hasData: true });
    
    // 生成日历
    this.generateCalendar();
    
    // 更新患者友好视图
    this.updatePatientFriendlyView();
  },

  // 生成示例EMG数据（日历格式）
  generateSampleEMGData() {
    const data = {};
    const now = new Date();
    
    // 生成过去60天的数据（覆盖2个月）
    for (let i = 59; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      
      // 随机决定是否有数据（70%概率有数据）
      if (Math.random() > 0.3) {
        const emgThreshold = (Math.random() * 40 + 30).toFixed(1); // 30-70之间的随机值
        const dateKey = this.formatDateKey(date);
        data[dateKey] = parseFloat(emgThreshold);
      }
    }
    
    return data;
  },

  // 格式化日期为（DD）
  formatDate1(dateKey) {
    const parts = dateKey.split('-');
    return parts[2];  // 返回日
  },

  // 格式化日期为（MM-DD）
  formatDate2(dateKey) {
    const parts = dateKey.split('-');
    return parts[1] + '-' + parts[2];  // 返回月-日
  },

  // 格式化日期为key（YYYY-MM-DD）
  formatDateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  },

  // 生成日历数据
  generateCalendar() {
    const { currentYear, currentMonth, emgData, selectedDate } = this.data;
    const calendarDays = [];
    
    // 获取当月第一天和最后一天
    const firstDay = new Date(currentYear, currentMonth - 1, 1);
    const lastDay = new Date(currentYear, currentMonth, 0);
    
    // 获取第一天是星期几（0=周日）
    const firstDayWeek = firstDay.getDay();
    
    // 添加上个月的末尾几天
    const prevMonth = new Date(currentYear, currentMonth - 2, 0);
    for (let i = firstDayWeek - 1; i >= 0; i--) {
      const day = prevMonth.getDate() - i;
      const date = new Date(currentYear, currentMonth - 2, day);
      const dateKey = this.formatDateKey(date);
      calendarDays.push({
        date: day,
        fullDate: date,
        dateKey: dateKey,
        isCurrentMonth: false,
        isToday: false,
        hasEMGData: emgData[dateKey] !== undefined,
        emgValue: emgData[dateKey] || null,
        isSelected: selectedDate === dateKey
      });
    }
    
    // 添加当月的所有天
    for (let day = 1; day <= lastDay.getDate(); day++) {
      const date = new Date(currentYear, currentMonth - 1, day);
      const today = new Date();
      const isToday = date.toDateString() === today.toDateString();
      const dateKey = this.formatDateKey(date);
      
      calendarDays.push({
        date: day,
        fullDate: date,
        dateKey: dateKey,
        isCurrentMonth: true,
        isToday: isToday,
        hasEMGData: emgData[dateKey] !== undefined,
        emgValue: emgData[dateKey] || null,
        isSelected: selectedDate === dateKey
      });
    }
    
    // 添加下个月的开头几天（补齐6行，共42天）
    const remainingDays = 42 - calendarDays.length;
    for (let day = 1; day <= remainingDays; day++) {
      const date = new Date(currentYear, currentMonth, day);
      const dateKey = this.formatDateKey(date);
      calendarDays.push({
        date: day,
        fullDate: date,
        dateKey: dateKey,
        isCurrentMonth: false,
        isToday: false,
        hasEMGData: emgData[dateKey] !== undefined,
        emgValue: emgData[dateKey] || null,
        isSelected: selectedDate === dateKey
      });
    }
    
    this.setData({ calendarDays: calendarDays });
  },

 

  // 更新患者友好视图
  updatePatientFriendlyView() {
    const { emgData } = this.data;
    const today = this.formatDateKey(new Date());
    const yesterday = this.formatDateKey(new Date(Date.now() - 24 * 60 * 60 * 1000));
    
    // 计算今日状态
    const todayValue = emgData[today] || 0;
    const yesterdayValue = emgData[yesterday] || 0;
    const improvement = todayValue - yesterdayValue;
    
    // 确定状态
    let status = 'stable';
    let message = '保持稳定';
    
    if (improvement > 2) {
      status = 'good';
      message = '太棒了！比昨天进步了！';
    } else if (improvement < -2) {
      status = 'attention';
      message = '今天有点下降，明天加油！';
    } else if (todayValue > 45) {
      status = 'good';
      message = '状态很好，继续努力！';
    }
    
    // 计算周进度
    const currentDate = new Date();
    const weeklyValues = this.getWeekData(currentDate).values;
    const weeklyAverage = weeklyValues.length > 0 ? 
      weeklyValues.reduce((sum, val) => sum + Number(val), 0) / 7 : 0; 
      const completion = Math.min(100, (weeklyAverage / 50) * 100);
    
    
    // 计算统计数据
    const totalTrainingDays = Object.keys(emgData).length;
    const bestValue = totalTrainingDays > 0 ? 
      Math.max(...Object.values(emgData).map(val => parseFloat(val)).filter(val => !isNaN(val))).toFixed(1) : '0';
  
   
    // 更新数据
    this.setData({
      currentStatus: {
        todayValue: todayValue,
        yesterdayValue: yesterdayValue,
        improvement: improvement,
        status: status,
        message: message
      },
      weeklyProgress: {
        average: weeklyAverage.toFixed(1),
        target: 50, // 用户设置的目标值：默认50
        completion: completion // 周进度完成率
      },
      totalTrainingDays: totalTrainingDays,
      bestValue: bestValue
    });
    // console.log("本周平均:",this.data.weeklyProgress)
    // 检查成就
    this.checkAchievements();
  },



  // 检查成就
  checkAchievements() {
    const { emgData } = this.data;
    const achievements = [];
    
    // 检查连续达标天数
    let consecutiveDays = 0;
    const today = new Date();
    for (let i = 0; i < 30; i++) {
      const date = new Date(today);
      date.setDate(date.getDate() - i);
      const dateKey = this.formatDateKey(date);
      if (emgData[dateKey] && emgData[dateKey] >= 45) {
        consecutiveDays++;
      } else {
        break;
      }
    }
    
    if (consecutiveDays >= 7) {
      achievements.push({
        icon: '🏆',
        title: '连续达标',
        desc: `连续${consecutiveDays}天达到目标`
      });
    }
    
    // 检查最高值
    const values = Object.values(emgData);
    if (values.length > 0) {
      const numericValues = values.map(val => parseFloat(val)).filter(val => !isNaN(val));
      if (numericValues.length > 0) {
        const maxValue = Math.max(...numericValues);
        achievements.push({
          icon: '🌟',
          title: '最高记录',
          desc: `${maxValue.toFixed(1)}μV`
        });
      }
    }
    
    // 检查进步最大的一天
    const dates = Object.keys(emgData).sort();
    let maxImprovement = 0;
    let bestDay = '';
    for (let i = 1; i < dates.length; i++) {
      const currentValue = parseFloat(emgData[dates[i]]);
      const previousValue = parseFloat(emgData[dates[i-1]]);
      
      if (!isNaN(currentValue) && !isNaN(previousValue)) {
        const improvement = currentValue - previousValue;
        if (improvement > maxImprovement) {
          maxImprovement = improvement;
          bestDay = dates[i];
        }
      }
    }
    
    if (maxImprovement > 5) {
      achievements.push({
        icon: '💪',
        title: '最大进步',
        desc: `${bestDay} 进步${maxImprovement.toFixed(1)}`
      });
    }
    
    this.setData({ achievements: achievements });
  },


  // 选择周
  selectWeek() {
    const { selectedDateInfo } = this.data;
    
    if (!selectedDateInfo) {
      wx.showToast({
        title: '请先选择一个日期',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 基于选中日期计算该周的所有信息
    const selectedDate = selectedDateInfo.date;
    const weekData = this.getWeekData(selectedDate);
    
    // 更新selectedDateInfo，添加周信息
    this.setData({
      selectedDateInfo: {
        ...selectedDateInfo,
        weekStr: weekData.weekStr,
        weekStart: weekData.weekStart,
        weekEnd: weekData.weekEnd,
        weekDates: weekData.weekDates
      }
    });
    // console.log("所选日期的周数据:",this.data.selectedDateInfo.weekDates)
    
    // 延迟显示Toast，确保setData完成，使用简洁文本
    setTimeout(() => {
      wx.showToast({
        title: `${weekData.weekStr}`,
        icon: 'success',
        duration: 2000
      });
    }, 500);
    
    // 显示周数据弹窗图表
    this.showWeekChartModal();
  },


  // 选择月
  selectMonth() {
    const { selectedDateInfo } = this.data;
    
    if (!selectedDateInfo) {
      wx.showToast({
        title: '请先选择一个日期',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 基于选中日期计算该月
    const selectedDate = selectedDateInfo.date;
    const monthStr = `${selectedDate.getMonth() + 1}月`;
    
    // 计算该月的所有日期
    const monthDates = [];
    const firstDay = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1);
    const lastDay = new Date(selectedDate.getFullYear(), selectedDate.getMonth() + 1, 0);
    
    for (let day = 1; day <= lastDay.getDate(); day++) {
      const date = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), day);
      monthDates.push(this.formatDateKey(date));
    }
    
    // 更新selectedDateInfo，添加月信息
    this.setData({
      selectedDateInfo: {
        ...selectedDateInfo,
        monthStr: monthStr,
        monthFirstDay: firstDay,
        monthLastDay: lastDay,
        monthDates: monthDates
      }
    });
    
    // 延迟显示Toast，确保setData完成，使用简洁文本
    setTimeout(() => {
      wx.showToast({
        title: `${monthStr}`,
        icon: 'success',
        duration: 2000
      });
    }, 500);
    
    // 显示月数据弹窗图表
    this.showMonthChartModal();
  },






  // 日历月份切换 - 上一月
  goToPreviousMonth() {
    let { currentYear, currentMonth } = this.data;
    currentMonth--;
    if (currentMonth < 1) {
      currentMonth = 12;
      currentYear--;
    }
    this.setData({ currentYear, currentMonth });
    this.generateCalendar();
  },

  // 日历月份切换 - 下一月
  goToNextMonth() {
    let { currentYear, currentMonth } = this.data;
    currentMonth++;
    if (currentMonth > 12) {
      currentMonth = 1;
      currentYear++;
    }
    this.setData({ currentYear, currentMonth });
    this.generateCalendar();
  },

  // 选择日期
  selectDate(e) {
    const dateKey = e.currentTarget.dataset.datekey;
    const selectedDate = this.data.selectedDate;
    
    // 验证数据完整性
    if (!dateKey) {
      console.error('selectDate: dateKey为空');
      wx.showToast({
        title: '选择失败：日期信息不完整',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 获取EMG数据，确保数据完整性
    const emgData = this.data.emgData || {};
    const emgValue = emgData[dateKey];
    const hasEMGData = emgValue !== undefined && emgValue !== null && !isNaN(emgValue);
    
    console.log('selectDate调试信息:', {
      dateKey: dateKey,
      emgValue: emgValue,
      hasEMGData: hasEMGData,
      emgDataType: typeof emgValue,
      currentSelectedDate: selectedDate
    });
    
    // 如果点击的是已选中的日期，则取消选择
    if (selectedDate === dateKey) {
      this.setData({ 
        selectedDate: null,
        selectedDateInfo: null
      });
      
      // 延迟显示Toast，确保setData完成
      setTimeout(() => {
        wx.showToast({
          title: '已取消选择',
          icon: 'success',
          duration: 1000
        });
      }, 50);
      
    } else {
      // 选择新日期
      const date = new Date(dateKey);
      
      // 验证日期有效性
      if (isNaN(date.getTime())) {
        console.error('selectDate: 无效的日期', dateKey);
        wx.showToast({
          title: '选择失败：无效日期',
          icon: 'none',
          duration: 2000
        });
        return;
      }
      
      const selectedDateInfo = {
        dateKey: dateKey,
        date: date,
        year: date.getFullYear(),
        month: date.getMonth() + 1,
        day: date.getDate(),
        weekDay: date.getDay(), // 0=周日, 1=周一...
        emgValue: hasEMGData ? parseFloat(emgValue) : null,
        hasEMGData: hasEMGData
      };
      
      this.setData({ 
        selectedDate: dateKey,
        selectedDateInfo: selectedDateInfo
      });
      
      // // 延迟显示Toast，确保setData完成，并构建完整信息
      // setTimeout(() => {
      //   let toastTitle = '';
      //   if (hasEMGData) {
      //     const formattedValue = parseFloat(emgValue).toFixed(1);
      //     toastTitle = `已选择 ${dateKey} (${formattedValue}μV)`;
      //   } else {
      //     toastTitle = `已选择 ${dateKey} (无数据)`;
      //   }
      //   wx.showToast({
      //     title: toastTitle,
      //     icon: 'success',
      //     duration: 2000
      //   });
      // }, 50);
    }
    
    // 重新生成日历以更新选中状态
    this.generateCalendar();
  },


  // 刷新数据
  refreshData() {
    this.loadEMGData();
  },

  // 获取周数据
  getWeekData(selectedDate) {
    const { emgData } = this.data;
    const dates = [];
    const values = [];
    
    // 计算该周的开始日期（周日）
    const weekStart = new Date(selectedDate);
    weekStart.setDate(selectedDate.getDate() - selectedDate.getDay());
    
    // 计算该周的结束日期（周六）
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    
    // 生成周字符串
    const weekStr = `${weekStart.getMonth()+1}/${weekStart.getDate()}-${weekEnd.getMonth()+1}/${weekEnd.getDate()}`;
    
    // 生成完整的周日期数组
    const weekDates = [];
    for (let i = 0; i < 7; i++) {
      const date = new Date(weekStart);
      date.setDate(weekStart.getDate() + i);
      weekDates.push(this.formatDateKey(date));
    }
    
    // 获取该周7天的数据（仅包含有EMG数据的日期）
    for (let i = 0; i < 7; i++) {
      const date = new Date(weekStart);
      date.setDate(weekStart.getDate() + i);
      const dateKey = this.formatDateKey(date);
      
      if (emgData[dateKey] !== undefined) {
        const dateStr = `${date.getMonth()+1}/${date.getDate()}`;
        dates.push(dateStr);
        values.push(emgData[dateKey]);
      }
    }
    
    return { 
      dates, 
      values, 
      weekStart, 
      weekEnd, 
      weekStr, 
      weekDates 
    };
  },

  // 获取月数据
  getMonthData(selectedDate) {
    const { emgData } = this.data;
    const dates = [];
    const values = [];
    
    // 获取该月所有有数据的日期
    const firstDay = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1);
    const lastDay = new Date(selectedDate.getFullYear(), selectedDate.getMonth() + 1, 0);
    
    for (let day = 1; day <= lastDay.getDate(); day++) {
      const date = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), day);
      const dateKey = this.formatDateKey(date);
      
      if (emgData[dateKey] !== undefined) {
        const dateStr = `${date.getMonth()+1}/${date.getDate()}`;
        dates.push(dateStr);
        values.push(emgData[dateKey]);
      }
    }
    
    return { dates, values };
  },

  // 关闭图表弹窗
  closeChartModal() {
    this.setData({
      showChartModal: false,
      modalInitialized: false,
      chartModalData: {
        title: '',
        dates: [],
        values: [],
        type: 'line',
        average: 0,
        max: 0,
        min: 0
      },
      // 重置交互状态
      isScaling: false,
      isPanning: false,
      chartScale: 1,
      chartPanX: 0,
      chartPanY: 0
    });
  },

  // 等待弹窗DOM渲染完成
  waitForModalRender(callback) {
    let retryCount = 0;
    const maxRetries = 20; // 最多重试20次，即1秒
    
    const checkModal = () => {
    const query = wx.createSelectorQuery();
      query.select('.canvas-container').boundingClientRect((rect) => {
        if (rect && rect.width > 0 && rect.height > 0) {
          // DOM已渲染完成，延迟一点时间确保动画完成
          setTimeout(() => {
            callback();
          }, 300); // 增加延迟时间
        } else if (retryCount < maxRetries) {
          retryCount++;
          // 继续等待
          setTimeout(checkModal, 50);
        } else {
          console.error('弹窗DOM渲染超时，强制绘制图表');
          callback();
        }
      }).exec();
    };
    
    // 开始检查
    setTimeout(checkModal, 150); // 增加初始延迟
  },

  // 计算统计数据
  calculateStats(values) {
    if (!values || values.length === 0) {
      return { average: 0, max: 0, min: 0 };
    }
    
    // 确保所有值都是数字类型，过滤掉无效值
    const numericValues = values
      .map(val => {
        const numVal = parseFloat(val);
        return isNaN(numVal) ? null : numVal;
      })
      .filter(val => val !== null);
    
    if (numericValues.length === 0) {
      return { average: 0, max: 0, min: 0 };
    }
    
    const sum = numericValues.reduce((acc, val) => acc + val, 0);
    const average = sum / numericValues.length;
    const max = Math.max(...numericValues);
    const min = Math.min(...numericValues);
    
    return {
      average: parseFloat(average.toFixed(1)),
      max: parseFloat(max.toFixed(1)),
      min: parseFloat(min.toFixed(1))
    };
  },

  // 测试Canvas弹窗（简化版本）
  testEChartsModal() {
    console.log('测试Canvas弹窗');
    
    // 生成测试数据
    const testDates = ['1/1', '1/2', '1/3', '1/4', '1/5', '1/6', '1/7'];
    const testValues = [45, 52, 48, 55, 50, 47, 53];
    const stats = this.calculateStats(testValues);
    
    console.log('测试数据:', { testDates, testValues, stats });
    
    this.setData({
      showChartModal: true,
      modalInitialized: false,
      chartModalData: {
        title: '测试图表',
        dates: testDates,
        values: testValues,
        type: 'line',
        average: stats.average,
        max: stats.max,
        min: stats.min
      }
    });
    
    // 等待弹窗渲染完成后绘制图表
    this.waitForModalRender(() => {
      this.drawNativeChart(testDates, testValues, '测试图表');
    });
  },

  // 获取设备适配参数
  getDeviceAdaptationParams() {
    try {
      // 使用新的API获取设备信息
      const deviceInfo = wx.getDeviceInfo();
      const windowInfo = wx.getWindowInfo();
      const systemInfo = wx.getSystemInfoSync();
      
      const windowWidth = windowInfo?.windowWidth || 375;
      const windowHeight = windowInfo?.windowHeight || 667;
      const pixelRatio = deviceInfo?.pixelRatio || windowInfo?.pixelRatio || 1;
      
      // 检测iOS设备
      const isIOS = systemInfo.platform === 'ios';
      const isIPhoneX = systemInfo.model && systemInfo.model.includes('iPhone X');
      
      // 计算适配比例（以iPhone 6为基准：375px）
      const scaleRatio = windowWidth / 375;
      
      // iOS特殊处理 - 根据文档优化
      const iosAdjustment = isIOS ? {
        // iOS设备需要更保守的DPR设置，避免Canvas位置偏移
        adjustedPixelRatio: Math.min(pixelRatio, 1.2), // 进一步限制最大DPR为1.2
        // iOS设备不需要坐标偏移，避免位置问题
        coordinateOffset: { x: 0, y: 0 },
        // iOS安全区域适配
        safeAreaInsets: {
          top: systemInfo.safeArea?.top || 0,
          bottom: systemInfo.safeArea?.bottom || 0,
          left: systemInfo.safeArea?.left || 0,
          right: systemInfo.safeArea?.right || 0
        }
      } : {
        adjustedPixelRatio: Math.min(pixelRatio, 1.5), // 非iOS设备也限制最大DPR为1.5
        coordinateOffset: { x: 0, y: 0 },
        safeAreaInsets: { top: 0, bottom: 0, left: 0, right: 0 }
      };
      
      console.log('设备适配参数:', {
        windowWidth,
        windowHeight,
        pixelRatio,
        scaleRatio,
        isIOS,
        isIPhoneX,
        iosAdjustment,
        deviceInfo,
        windowInfo,
        systemInfo
      });
      
      return {
        windowWidth,
        windowHeight,
        pixelRatio: iosAdjustment.adjustedPixelRatio,
        scaleRatio,
        isIOS,
        isIPhoneX,
        coordinateOffset: iosAdjustment.coordinateOffset,
        safeAreaInsets: iosAdjustment.safeAreaInsets
      };
    } catch (error) {
      console.error('获取设备适配参数失败:', error);
      // 返回默认参数
      return {
        windowWidth: 375,
        windowHeight: 667,
        pixelRatio: 1,
        scaleRatio: 1,
        isIOS: false,
        isIPhoneX: false,
        coordinateOffset: { x: 0, y: 0 },
        safeAreaInsets: { top: 0, bottom: 0, left: 0, right: 0 }
      };
    }
  },

  // 获取Canvas备用尺寸 - iOS优化版本
  getFallbackCanvasSize() {
    try {
      // 使用新的API获取窗口信息
      const windowInfo = wx.getWindowInfo();
      const systemInfo = wx.getSystemInfoSync();
      const windowWidth = windowInfo?.windowWidth || 375; // 默认iPhone 6宽度
      const windowHeight = windowInfo?.windowHeight || 667; // 默认iPhone 6高度
      
      // 检测iOS设备
      const isIOS = systemInfo.platform === 'ios';
      const isIPhoneX = systemInfo.model && systemInfo.model.includes('iPhone X');
      
      // 计算弹窗的理论尺寸（弹窗占屏幕的96%，Canvas容器占弹窗的90%）
      const modalWidth = windowWidth * 0.96;
      const modalHeight = windowHeight * 0.88;
      
      // iOS设备特殊处理
      let canvasContainerWidth, canvasContainerHeight;
      if (isIOS) {
        // iOS设备：考虑安全区域和状态栏
        canvasContainerWidth = modalWidth * 0.9; // 减去padding
        canvasContainerHeight = isIPhoneX ? 500 : 550; // iPhone X系列需要调整高度
      } else {
        // 非iOS设备：标准处理
        canvasContainerWidth = modalWidth * 0.9; // 减去padding
        canvasContainerHeight = 550; // CSS中设置的固定高度
      }
      
      
      return {
        width: canvasContainerWidth,
        height: canvasContainerHeight
      };
    } catch (error) {
      console.error('获取备用Canvas尺寸失败:', error);
      // 返回默认尺寸
      return {
        width: 600,
        height: 400
      };
    }
  },

  // Canvas位置修复函数 - 修复层级问题
  fixIOSCanvasPosition(canvas, containerWidth, containerHeight) {
    try {
      const systemInfo = wx.getSystemInfoSync();
      const isIOS = systemInfo.platform === 'ios';
      
      
      // 检查Canvas和style属性是否存在
      if (!canvas) {
        console.warn('Canvas元素不存在');
        return;
      }
      
      if (!canvas.style) {
        console.warn('Canvas style属性不存在，尝试初始化');
        // 在微信小程序中，Canvas的style属性可能不存在，我们通过CSS类来控制
        canvas.className = 'modal-chart-canvas';
        console.log('已设置Canvas className');
        return;
      }
      
      // 关键修复：确保Canvas在容器内正确定位
      canvas.style.position = 'absolute';
      canvas.style.left = '0px';
      canvas.style.top = '0px';
      canvas.style.margin = '0';
      canvas.style.padding = '0';
      
      // 确保Canvas尺寸正确
      canvas.style.width = '100%';
      canvas.style.height = '100%';
      canvas.style.maxWidth = '100%';
      canvas.style.maxHeight = '100%';
      
      // 设置z-index确保层级正确
      canvas.style.zIndex = '2';
      
      // 应用硬件加速
      canvas.style.transform = 'translateZ(0)';
      canvas.style.webkitTransform = 'translateZ(0)';
      
    } catch (error) {
      console.error('Canvas位置修复失败:', error);
    }
  },

  // iOS Canvas初始化优化
  initializeIOSCanvas(canvas, ctx, containerWidth, containerHeight) {
    try {
      const systemInfo = wx.getSystemInfoSync();
      const isIOS = systemInfo.platform === 'ios';
      
      if (isIOS) {
        console.log('初始化iOS Canvas');
        
        // 保存原始变换状态
        ctx.save();
        
        // 重置变换矩阵
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        
        // 设置Canvas原点位置
        ctx.translate(0, 0);
        
        // 应用iOS特定的渲染优化
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        
      }
    } catch (error) {
      console.error('iOS Canvas初始化失败:', error);
    }
  },

  // Canvas位置验证函数 - 微信小程序兼容版本
  validateCanvasPosition(canvas, containerWidth, containerHeight) {
    try {
      // 微信小程序中Canvas没有getBoundingClientRect方法，使用其他方式验证
      console.log('Canvas位置验证:', {
        canvasDimensions: {
          width: canvas.width,
          height: canvas.height,
          styleWidth: canvas.style?.width,
          styleHeight: canvas.style?.height
        },
        containerDimensions: {
          width: containerWidth,
          height: containerHeight
        },
        canvasProperties: {
          nodeType: canvas.nodeType,
          tagName: canvas.tagName,
          id: canvas.id,
          className: canvas.className
        }
      });
      
      // 检查Canvas基本属性
      if (canvas && canvas.width > 0 && canvas.height > 0) {
        console.log('Canvas基本属性验证通过');
        
        // 检查Canvas尺寸是否合理 - 更严格的验证
        const maxAllowedWidth = containerWidth * 3; // 允许最大3倍容器宽度
        const maxAllowedHeight = containerHeight * 3; // 允许最大3倍容器高度
        
        if (canvas.width <= maxAllowedWidth && canvas.height <= maxAllowedHeight) {
          return true;
        } else {
          console.warn('Canvas尺寸异常:', {
            canvasWidth: canvas.width,
            canvasHeight: canvas.height,
            containerWidth,
            containerHeight,
            maxAllowedWidth,
            maxAllowedHeight,
            widthRatio: (canvas.width / containerWidth).toFixed(2),
            heightRatio: (canvas.height / containerHeight).toFixed(2)
          });
          return false;
        }
      } else {
        console.warn('Canvas基本属性验证失败');
        return false;
      }
    } catch (error) {
      console.error('Canvas位置验证出错:', error);
      return false;
    }
  },

  // Canvas状态检查函数 - 微信小程序专用
  checkCanvasStatus(canvas) {
    try {
      console.log('Canvas状态检查:', {
        canvasExists: !!canvas,
        canvasType: typeof canvas,
        canvasConstructor: canvas?.constructor?.name,
        canvasMethods: {
          hasGetContext: typeof canvas?.getContext === 'function',
          hasWidth: 'width' in canvas,
          hasHeight: 'height' in canvas,
          hasStyle: 'style' in canvas
        },
        canvasValues: {
          width: canvas?.width,
          height: canvas?.height,
          styleWidth: canvas?.style?.width,
          styleHeight: canvas?.style?.height
        }
      });
      
      // 检查Canvas是否可用
      if (canvas && typeof canvas.getContext === 'function') {
        console.log('Canvas状态检查通过');
        return true;
      } else {
        console.warn('Canvas状态检查失败');
        return false;
      }
    } catch (error) {
      console.error('Canvas状态检查出错:', error);
      return false;
    }
  },

  // 智能像素比计算函数
  calculateOptimalPixelRatio(originalPixelRatio, containerWidth, containerHeight, isIOS = false) {
    try {
      // 基础限制
      let maxPixelRatio = isIOS ? 1.5 : 2;
      
      // 根据容器尺寸动态调整
      if (containerWidth < 200 || containerHeight < 200) {
        // 小容器需要更低的像素比
        maxPixelRatio = Math.min(maxPixelRatio, 1.2);
      } else if (containerWidth > 400 || containerHeight > 400) {
        // 大容器可以使用稍高的像素比
        maxPixelRatio = Math.min(maxPixelRatio, 2.5);
      }
      
      const optimalPixelRatio = Math.min(originalPixelRatio, maxPixelRatio);
      
      console.log('智能像素比计算:', {
        原始像素比: originalPixelRatio,
        容器宽度: containerWidth,
        容器高度: containerHeight,
        最大允许像素比: maxPixelRatio,
        最终像素比: optimalPixelRatio,
        是iOS设备: isIOS
      });
      
      return optimalPixelRatio;
    } catch (error) {
      console.error('智能像素比计算失败:', error);
      return Math.min(originalPixelRatio, 1.5);
    }
  },

  // Canvas尺寸自动修复函数
  autoFixCanvasSize(canvas, targetWidth, targetHeight, pixelRatio) {
    try {
      
      // 计算合理的像素比
      const maxAllowedRatio = Math.min(2, Math.max(1, Math.min(
        (targetWidth * 3) / canvas.width,
        (targetHeight * 3) / canvas.height
      )));
      
      const fixedPixelRatio = Math.min(pixelRatio, maxAllowedRatio);
      
      // 重新设置Canvas尺寸
      canvas.width = targetWidth * fixedPixelRatio;
      canvas.height = targetHeight * fixedPixelRatio;
      canvas.style.width = targetWidth + 'px';
      canvas.style.height = targetHeight + 'px';
      
      
      return true;
    } catch (error) {
      console.error('Canvas尺寸自动修复失败:', error);
      return false;
    }
  },

  // 微信小程序Canvas位置查询函数
  queryCanvasPosition(canvasId) {
    return new Promise((resolve, reject) => {
      const query = wx.createSelectorQuery();
      query.select(`#${canvasId}`)
        .fields({
          node: true,
          size: true,
          rect: true,
          scrollOffset: true
        })
        .exec((res) => {
          if (res && res[0]) {
            console.log('Canvas位置查询结果:', res[0]);
            resolve(res[0]);
          } else {
            console.warn('Canvas位置查询失败');
            reject(new Error('Canvas位置查询失败'));
          }
        });
    });
  },

  // Canvas渲染位置检测函数 - 根据文档优化
  detectCanvasRenderPosition() {
    try {
      
      // 检查Canvas是否在正确位置渲染
      const query = wx.createSelectorQuery();
      query.select('.canvas-container')
        .fields({ 
          node: true, 
          size: true, 
          rect: true,
          computedStyle: ['position', 'overflow', 'zIndex', 'display'] 
        })
        .exec((res) => {
          if (res && res[0]) {
            const containerInfo = res[0];
            const styles = containerInfo.computedStyle || {};
            
            
            // 检查Canvas是否在容器内正确渲染
            if (containerInfo.width <= 0 || containerInfo.height <= 0) {
              console.warn('❌ Canvas容器尺寸异常');
            }
            
            // 检查容器定位是否正确（computedStyle可能返回undefined，但CSS已正确设置）
            if (styles.position !== 'relative' || styles.overflow !== 'hidden') {
              console.log('ℹ️ Canvas容器定位通过CSS设置（computedStyle可能未正确获取）');
            }
          } else {
            console.warn('❌ 无法获取Canvas容器信息');
          }
        });
    } catch (error) {
      console.error('Canvas渲染位置检测失败:', error);
    }
  },

  // 原生Canvas绘制图表（完全兼容弹窗）
  drawNativeChart(dates, values, title) {
    
    const query = wx.createSelectorQuery();
    query.select('#modalChart')
      .fields({ node: true, size: true })
      .exec((res) => {
        console.log('Canvas查询结果:', res);
        
        if (!res || !res[0] || !res[0].node) {
          console.error('Canvas节点未找到');
          return;
        }
        
        const canvas = res[0].node;
        const ctx = canvas.getContext('2d');
        
        if (!ctx) {
          console.error('无法获取Canvas上下文');
          return;
        }
        
        // 获取容器尺寸 - 确保获取到有效尺寸
        let containerWidth = res[0]?.width;
        let containerHeight = res[0]?.height;
        
        // 如果获取不到有效尺寸，使用备用尺寸计算
        if (!containerWidth || containerWidth <= 0 || containerWidth === undefined) {
          const fallbackSize = this.getFallbackCanvasSize();
          containerWidth = fallbackSize.width;
          console.warn('容器宽度获取失败，使用备用尺寸:', containerWidth);
        }
        
        if (!containerHeight || containerHeight <= 0 || containerHeight === undefined) {
          const fallbackSize = this.getFallbackCanvasSize();
          containerHeight = fallbackSize.height;
          console.warn('容器高度获取失败，使用备用尺寸:', containerHeight);
        }
        
        // 考虑容器的padding，调整实际可用尺寸
        const padding = 16; // 16rpx转换为px（假设1rpx = 0.5px）
        const actualWidth = containerWidth - padding * 2;
        const actualHeight = containerHeight - padding * 2;
        
        
        // 设置Canvas尺寸 - 智能像素比优化
        const canvasDeviceParams = this.getDeviceAdaptationParams();
        const { pixelRatio: originalPixelRatio, isIOS, coordinateOffset } = canvasDeviceParams;
        
        // 使用智能像素比计算
        const optimalPixelRatio = this.calculateOptimalPixelRatio(
          originalPixelRatio, 
          actualWidth, 
          actualHeight, 
          isIOS
        );
        
        console.log('Canvas设置参数:', {
          canvasDeviceParams,
          containerWidth,
          containerHeight,
          actualWidth,
          actualHeight,
          原始像素比: originalPixelRatio,
          优化像素比: optimalPixelRatio,
          isIOS
        });
        
        // 使用优化后的像素比设置Canvas尺寸
        try {
          if (isIOS) {
            // iOS设备：使用优化后的像素比
            canvas.width = actualWidth * optimalPixelRatio;
            canvas.height = actualHeight * optimalPixelRatio;
            
            // 设置Canvas样式尺寸（显示尺寸）
            canvas.style.width = actualWidth + 'px';
            canvas.style.height = actualHeight + 'px';
            
            // 应用坐标偏移（如果需要）
            if (coordinateOffset.x !== 0 || coordinateOffset.y !== 0) {
              ctx.translate(coordinateOffset.x, coordinateOffset.y);
            }
            
            // 应用像素比缩放
            ctx.scale(optimalPixelRatio, optimalPixelRatio);
          } else {
            // 非iOS设备：使用优化后的像素比
            canvas.width = actualWidth * optimalPixelRatio;
            canvas.height = actualHeight * optimalPixelRatio;
            ctx.scale(optimalPixelRatio, optimalPixelRatio);
          }
        } catch (error) {
          console.error('Canvas尺寸设置失败:', error);
          // 使用备用尺寸设置
          canvas.width = actualWidth;
          canvas.height = actualHeight;
          canvas.style.width = actualWidth + 'px';
          canvas.style.height = actualHeight + 'px';
        }
        
        // 验证Canvas尺寸设置是否成功
        if (canvas.width === 0 || canvas.height === 0) {
          console.error('Canvas尺寸设置失败:', {
            canvasWidth: canvas.width,
            canvasHeight: canvas.height
          });
          return;
        }
        
        
        // iOS Canvas位置修复
        this.fixIOSCanvasPosition(canvas, actualWidth, actualHeight);
        
        // iOS Canvas初始化优化
        this.initializeIOSCanvas(canvas, ctx, actualWidth, actualHeight);
        
        // 清除画布 - 确保使用正确的尺寸
        ctx.clearRect(0, 0, actualWidth, actualHeight);
        
        // 验证绘制参数
        if (!actualWidth || !actualHeight || actualWidth <= 0 || actualHeight <= 0) {
          console.error('无效的绘制参数:', { actualWidth, actualHeight });
          return;
        }
        
        // 验证Canvas是否可用
        if (!canvas || !ctx) {
          console.error('Canvas或Context不可用');
          return;
        }
        
        // 获取设备适配参数
        const renderDeviceParams = this.getDeviceAdaptationParams();
        
        // 绘制图表
        this.renderChart(ctx, dates, values, title, actualWidth, actualHeight, renderDeviceParams);
        
        // 验证Canvas位置 - 使用微信小程序兼容的方法
        setTimeout(() => {
          const isValid = this.validateCanvasPosition(canvas, actualWidth, actualHeight);
          
          // 如果Canvas尺寸异常，尝试自动修复
          if (!isValid) {
            this.autoFixCanvasSize(canvas, actualWidth, actualHeight, optimalPixelRatio);
          }
          
          // 额外的Canvas状态检查
          this.checkCanvasStatus(canvas);
          
          // Canvas渲染位置检测
          this.detectCanvasRenderPosition();
          
          // 使用微信小程序API查询Canvas位置
          this.queryCanvasPosition('modalChart').then(positionInfo => {
            console.log('Canvas位置信息:', positionInfo);
          }).catch(error => {
            console.warn('Canvas位置查询失败:', error);
          });
        }, 100);
        
        // 添加Canvas调试信息 - 层级问题诊断
        console.log('Canvas调试信息:', {
          canvasElement: {
            width: canvas.width,
            height: canvas.height,
            styleWidth: canvas.style?.width,
            styleHeight: canvas.style?.height,
            position: canvas.style?.position,
            left: canvas.style?.left,
            top: canvas.style?.top,
            zIndex: canvas.style?.zIndex,
            nodeType: canvas.nodeType,
            tagName: canvas.tagName
          },
          containerInfo: {
            containerWidth,
            containerHeight,
            actualWidth,
            actualHeight,
            padding
          },
          deviceInfo: {
            isIOS: renderDeviceParams.isIOS,
            pixelRatio: renderDeviceParams.pixelRatio,
            coordinateOffset: renderDeviceParams.coordinateOffset
          },
          层级诊断: {
            建议: 'Canvas应该使用position: absolute, z-index: 2',
            容器建议: '容器应该使用position: relative, overflow: hidden, z-index: 1'
          }
        });
        
      });
  },

  // 渲染图表内容
  renderChart(ctx, dates, values, title, canvasWidth, canvasHeight, deviceParams = {}) {
    // 参数验证
    if (!dates || dates.length === 0 || !values || values.length === 0) {
      return;
    }
    
    // iOS坐标系转换处理
    if (deviceParams.isIOS) {
      console.log('应用iOS坐标系转换');
      
      // 保存当前状态
      ctx.save();
      
      // iOS设备需要特殊的坐标系处理
      // 1. 重置变换矩阵
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      
      // 2. 应用坐标偏移（如果需要）
      if (deviceParams.coordinateOffset) {
        ctx.translate(deviceParams.coordinateOffset.x, deviceParams.coordinateOffset.y);
      }
      
      // 3. 应用安全区域偏移
      if (deviceParams.safeAreaInsets) {
        const { top, left } = deviceParams.safeAreaInsets;
        ctx.translate(left, top);
      }
    }
    
    // 设置样式参数 - 优化padding以充分利用空间
    const padding = {
      top: Math.max(40, canvasHeight * 0.08),      // 顶部留白：容器高度的8%，最少40px
      right: Math.max(30, canvasWidth * 0.05),     // 右侧留白：容器宽度的5%，最少30px
      bottom: Math.max(50, canvasHeight * 0.12),   // 底部留白：容器高度的12%，最少50px
      left: Math.max(50, canvasWidth * 0.08)       // 左侧留白：容器宽度的8%，最少50px
    };
    
    const chartWidth = canvasWidth - padding.left - padding.right;
    const chartHeight = canvasHeight - padding.top - padding.bottom;
    
    // 应用设备适配参数
    const scaleRatio = deviceParams.scaleRatio || 1;
    
    
    // 绘制背景
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvasWidth, canvasHeight);
    
    // 绘制标题 - 自适应字体大小
    const titleFontSize = Math.max(14, Math.min(18, canvasWidth * 0.03));
    ctx.fillStyle = '#333333';
    ctx.font = `bold ${titleFontSize}px Arial, sans-serif`;
    ctx.textAlign = 'center';
    ctx.fillText(title, canvasWidth / 2, padding.top * 0.6);
    
    // 计算数据范围 - 确保所有值都是数字类型
    const numericValues = values.map(val => parseFloat(val)).filter(val => !isNaN(val));
    if (numericValues.length === 0) {
      return;
    }
    
    const minValue = Math.min(...numericValues);
    const maxValue = Math.max(...numericValues);
    const valueRange = maxValue - minValue;
    const valuePadding = valueRange * 0.1;
    
    // 绘制网格线
    ctx.strokeStyle = '#f0f0f0';
    ctx.lineWidth = 1;
    
    // 水平网格线
    for (let i = 0; i <= 5; i++) {
      const y = padding.top + (chartHeight / 5) * i;
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(padding.left + chartWidth, y);
      ctx.stroke();
    }
    
    // 垂直网格线
    const step = Math.max(1, Math.floor(dates.length / 8));
    for (let i = 0; i < dates.length; i += step) {
      const x = padding.left + (chartWidth / (dates.length - 1)) * i;
      ctx.beginPath();
      ctx.moveTo(x, padding.top);
      ctx.lineTo(x, padding.top + chartHeight);
      ctx.stroke();
    }
    
    // 绘制坐标轴
    ctx.strokeStyle = '#333333';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(padding.left, padding.top);
    ctx.lineTo(padding.left, padding.top + chartHeight);
    ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
    ctx.stroke();
    
    // 绘制Y轴标签 - 自适应字体大小
    const labelFontSize = Math.max(10, Math.min(14, canvasWidth * 0.025));
    ctx.fillStyle = '#666666';
    ctx.font = `${labelFontSize}px Arial, sans-serif`;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let i = 0; i <= 5; i++) {
      const value = minValue + (valueRange / 5) * i;
      const y = padding.top + chartHeight - (chartHeight / 5) * i;
      ctx.fillText(value.toFixed(1), padding.left - 8, y);
    }
    
    // 绘制X轴标签 - 自适应字体大小
    const xLabelFontSize = Math.max(9, Math.min(12, canvasWidth * 0.02));
    ctx.textAlign = 'center';
    ctx.font = `${xLabelFontSize}px Arial, sans-serif`;
    for (let i = 0; i < dates.length; i += step) {
      const x = padding.left + (chartWidth / (dates.length - 1)) * i;
      ctx.fillText(dates[i], x, padding.top + chartHeight + 18);
    }
    
    // 绘制Y轴标题 - 自适应字体大小
    const axisTitleFontSize = Math.max(11, Math.min(13, canvasWidth * 0.022));
    ctx.save();
    ctx.translate(15, padding.top + chartHeight / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#666666';
    ctx.font = `${axisTitleFontSize}px Arial, sans-serif`;
    ctx.textAlign = 'center';
    ctx.fillText('EMG值(μV)', 0, 0);
    ctx.restore();
    
    // 绘制折线图
    if (values.length > 1) {
      // 绘制面积
      ctx.fillStyle = 'rgba(84, 112, 198, 0.2)';
      ctx.beginPath();
      ctx.moveTo(padding.left, padding.top + chartHeight);
      
      dates.forEach((date, index) => {
        const x = padding.left + (chartWidth / (dates.length - 1)) * index;
        const normalizedValue = (numericValues[index] - minValue + valuePadding) / (valueRange + 2 * valuePadding);
        const y = padding.top + chartHeight - normalizedValue * chartHeight;
        ctx.lineTo(x, y);
      });
      
      ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
      ctx.closePath();
      ctx.fill();
      
      // 绘制折线 - 自适应线条粗细
      const lineWidth = Math.max(2, Math.min(4, canvasWidth * 0.008));
      ctx.strokeStyle = '#5470c6';
      ctx.lineWidth = lineWidth;
      ctx.beginPath();
      
      dates.forEach((date, index) => {
        const x = padding.left + (chartWidth / (dates.length - 1)) * index;
        const normalizedValue = (numericValues[index] - minValue + valuePadding) / (valueRange + 2 * valuePadding);
        const y = padding.top + chartHeight - normalizedValue * chartHeight;
        
        if (index === 0) {
          ctx.moveTo(x, y);
      } else {
          ctx.lineTo(x, y);
        }
      });
      ctx.stroke();
      
      // 绘制数据点 - 自适应点大小
      const pointRadius = Math.max(4, Math.min(7, canvasWidth * 0.012));
      const pointBorderWidth = Math.max(1.5, Math.min(3, canvasWidth * 0.005));
      
      ctx.fillStyle = '#5470c6';
      dates.forEach((date, index) => {
        const x = padding.left + (chartWidth / (dates.length - 1)) * index;
        const normalizedValue = (numericValues[index] - minValue + valuePadding) / (valueRange + 2 * valuePadding);
        const y = padding.top + chartHeight - normalizedValue * chartHeight;
        
        ctx.beginPath();
        ctx.arc(x, y, pointRadius, 0, 2 * Math.PI);
        ctx.fill();
        
        // 数据点外圈
        ctx.strokeStyle = '#ffffff';
        ctx.lineWidth = pointBorderWidth;
        ctx.stroke();
      });
    }
    
    // 绘制说明文字 - 自适应字体大小
    // const tipFontSize = Math.max(9, Math.min(12, canvasWidth * 0.018));
    // ctx.fillStyle = '#999999';
    // ctx.font = `${tipFontSize}px Arial, sans-serif`;
    // ctx.textAlign = 'left';
    // ctx.fillText('支持手势缩放和拖拽', padding.left, canvasHeight - 8);
    
    // iOS状态恢复
    if (deviceParams.isIOS) {
      console.log('恢复iOS Canvas状态');
      ctx.restore();
    }
  },
  // ======== 动态移动 和 缩放 图表,查静置显示数据功能：待实现 ====
  // // 触摸手势处理
  // onTouchStart(e) {
  //   const touches = e.touches;
  //   if (touches.length === 2) {
  //     // 双指触摸开始
  //     const touch1 = touches[0];
  //     const touch2 = touches[1];
  //     const distance = this.getDistance(touch1, touch2);
      
  //     this.setData({
  //       isScaling: true,
  //       initialDistance: distance,
  //       initialScale: this.data.chartScale || 1
  //     });
      
  //     console.log('开始缩放，初始距离:', distance);
  //   } else if (touches.length === 1) {
  //     // 单指触摸开始
  //     const touch = touches[0];
  //     this.setData({
  //       isPanning: true,
  //       lastTouchX: touch.clientX,
  //       lastTouchY: touch.clientY,
  //       initialPanX: this.data.chartPanX || 0,
  //       initialPanY: this.data.chartPanY || 0
  //     });
      
  //     console.log('开始拖拽');
  //   }
  // },

  // onTouchMove(e) {
  //   const touches = e.touches;
    
  //   if (touches.length === 2 && this.data.isScaling) {
  //     // 双指缩放
  //     const touch1 = touches[0];
  //     const touch2 = touches[1];
  //     const currentDistance = this.getDistance(touch1, touch2);
      
  //     const scale = this.data.initialScale * (currentDistance / this.data.initialDistance);
  //     const clampedScale = Math.max(0.5, Math.min(3, scale)); // 限制缩放范围
      
  //     this.setData({
  //       chartScale: clampedScale
  //     });
      
  //     // 重新绘制图表
  //     if (this.data.chartModalData.dates.length > 0) {
  //       this.drawNativeChart(
  //         this.data.chartModalData.dates, 
  //         this.data.chartModalData.values, 
  //         this.data.chartModalData.title
  //       );
  //     }
      
  //   } else if (touches.length === 1 && this.data.isPanning) {
  //     // 单指拖拽
  //     const touch = touches[0];
  //     const deltaX = touch.clientX - this.data.lastTouchX;
  //     const deltaY = touch.clientY - this.data.lastTouchY;
      
  //     this.setData({
  //       chartPanX: this.data.initialPanX + deltaX,
  //       chartPanY: this.data.initialPanY + deltaY,
  //       lastTouchX: touch.clientX,
  //       lastTouchY: touch.clientY
  //     });
      
  //     // 重新绘制图表
  //     if (this.data.chartModalData.dates.length > 0) {
  //       this.drawNativeChart(
  //         this.data.chartModalData.dates, 
  //         this.data.chartModalData.values, 
  //         this.data.chartModalData.title
  //       );
  //     }
  //   }
  // },

  // onTouchEnd(e) {
  //   console.log('触摸结束');
  //   this.setData({
  //     isScaling: false,
  //     isPanning: false,
  //     initialDistance: 0,
  //     lastTouchX: 0,
  //     lastTouchY: 0
  //   });
  // },

  // // 计算两点间距离
  // getDistance(touch1, touch2) {
  //   const dx = touch1.clientX - touch2.clientX;
  //   const dy = touch1.clientY - touch2.clientY;
  //   return Math.sqrt(dx * dx + dy * dy);
  // },



 

  // 添加测试数据点（指定某天）
  addTestDataPoint() {
    
    
    wx.showActionSheet({
      itemList: dates,
      success: (res) => {
        // 根据选择的索引计算对应的日期
        const selectedIndex = res.tapIndex;
        const selectedDate = new Date();
        selectedDate.setDate(selectedDate.getDate() - (29 - selectedIndex));
        const dateKey = this.formatDateKey(selectedDate);
        this.addEMGDataForDate(dateKey);
      }
    });
  },

  // 为指定日期添加EMG数据
  addEMGDataForDate(dateKey) {
    const threshold = (Math.random() * 40 + 30).toFixed(1); // 30-70之间的随机值
    
    // 更新EMG数据
    const updatedEmgData = { ...this.data.emgData };
    updatedEmgData[dateKey] = parseFloat(threshold);
    
    this.setData({ emgData: updatedEmgData, hasData: true });
    
    // 重新生成日历
    this.generateCalendar();
    
    
    wx.showToast({
      title: `${dateKey}: ${threshold}`,
      icon: 'success',
      duration: 2000
    });
  },

  // 清空所有数据
  clearAllData() {
    wx.showModal({
      title: '确认清空',
      content: '确定要清空所有EMG数据吗？',
      success: (res) => {
        if (res.confirm) {
          this.setData({ 
            emgData: {},
            calendarDays: [],
            hasData: false,
            selectedDate: null,
            showChartModal: false
          });
          
          // 重新生成空日历
          this.generateCalendar();
          
          wx.showToast({
            title: '数据已清空',
            icon: 'success'
          });
    }}
  });
  },

  // ================================
  // 弹窗图表相关函数
  // ================================

  // 显示图表弹窗
  showChartModal(title, description, dates, values) {
    
    // 计算统计数据
    const stats = this.calculateStats(values);
    
    this.setData({
      showChartModal: true,
      modalInitialized: false,
      chartModalData: {
        title: title || 'EMG阈值趋势图',
        dates: dates || [],
        values: values || [],
        type: 'line',
        average: stats.average,
        max: stats.max,
        min: stats.min
      },
      // 重置交互状态
      isScaling: false,
      isPanning: false,
      chartScale: 1,
      chartPanX: 0,
      chartPanY: 0
    });
    
    // 等待弹窗DOM渲染完成后绘制图表
    this.waitForModalRender(() => {
      this.drawNativeChart(dates, values, title);
    });
  },



  // 显示周数据弹窗图表
  showWeekChartModal() {
    const { selectedDateInfo } = this.data;
    
    if (!selectedDateInfo || !selectedDateInfo.weekDates) {
      wx.showToast({
        title: '请先选择一个日期',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 获取该周的EMG数据
    const weekDates = selectedDateInfo.weekDates;
    const weekValues = weekDates.map(dateKey => {
      const value = this.data.emgData[dateKey];
      return value !== undefined ? value : null;
    });
   
    // 过滤掉没有数据的日期
    const validDates = [];
    const validValues = [];
    
    weekDates.forEach((dateKey, index) => {
      if (weekValues[index] !== null) {
        validDates.push(this.formatDate2(dateKey));
        validValues.push(weekValues[index]);
      }
    });
    
    if (validDates.length === 0) {
      wx.showToast({
        title: '该周暂无训练数据',
        icon: 'none',
        duration: 2000
      });
      return;
    }
 
  
    // 显示弹窗图表
    this.showChartModal(
      `周数据趋势 (${selectedDateInfo.weekStr})`,
      `显示${selectedDateInfo.weekStr}期间的EMG阈值变化趋势，共${validDates.length}个数据点`,
      validDates,
      validValues
    );
  },

  // 显示月数据弹窗图表
  showMonthChartModal() {
    const { selectedDateInfo } = this.data;
    
    if (!selectedDateInfo || !selectedDateInfo.monthDates) {
      wx.showToast({
        title: '请先选择一个日期',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 获取该月的EMG数据
    const monthDates = selectedDateInfo.monthDates;
    const monthValues = monthDates.map(dateKey => {
      const value = this.data.emgData[dateKey];
      return value !== undefined ? value : null;
    });
    
    // 过滤掉没有数据的日期
    const validDates = [];
    const validValues = [];
    
    monthDates.forEach((dateKey, index) => {
      if (monthValues[index] !== null) {
        validDates.push(this.formatDate1(dateKey));
        validValues.push(monthValues[index]);
      }
    });
    
    if (validDates.length === 0) {
      wx.showToast({
        title: '该月暂无训练数据',
        icon: 'none',
        duration: 2000
      });
      return;
    }
    
    // 显示弹窗图表
    this.showChartModal(
      `月数据趋势 (${selectedDateInfo.year}年${selectedDateInfo.monthStr})`,
      `显示${selectedDateInfo.monthStr}期间的EMG阈值变化趋势，共${validDates.length}个数据点`,
      validDates,
      validValues
    );
  }
});

