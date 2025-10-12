const app = getApp();

// 简化的原生canvas图表绘制
function drawChart(ctx, dates, values, canvasWidth, canvasHeight) {
  console.log('drawChart开始:', { 
    dates: dates ? dates.length : 0, 
    values: values ? values.length : 0, 
    canvasWidth, 
    canvasHeight,
    canvasWidthValid: !isNaN(canvasWidth) && canvasWidth > 0,
    canvasHeightValid: !isNaN(canvasHeight) && canvasHeight > 0
  });
  
  // 参数验证
  if (!dates || dates.length === 0 || !values || values.length === 0) {
    console.log('drawChart: 数据为空，跳过绘制');
    return;
  }
  
  if (!ctx) {
    console.error('drawChart: ctx为空');
    return;
  }
  
  if (isNaN(canvasWidth) || isNaN(canvasHeight) || canvasWidth <= 0 || canvasHeight <= 0) {
    console.error('drawChart: 尺寸无效', { canvasWidth, canvasHeight });
    return;
  }

  // 清除画布
  ctx.clearRect(0, 0, canvasWidth, canvasHeight);
  
  // 设置样式 - 根据canvas大小调整padding
  const padding = Math.max(30, canvasWidth * 0.08); // 至少30px，或canvas宽度的8%
  const chartWidth = canvasWidth - 2 * padding;
  const chartHeight = canvasHeight - 2 * padding;
  
  console.log('drawChart尺寸:', { 
    canvasWidth, 
    canvasHeight, 
    padding, 
    chartWidth, 
    chartHeight,
    chartWidthValid: chartWidth > 0,
    chartHeightValid: chartHeight > 0
  });
  
  // 确保图表尺寸有效
  if (chartWidth <= 0 || chartHeight <= 0) {
    console.error('drawChart: 计算出的图表尺寸无效', { chartWidth, chartHeight });
    return;
  }
  
  // 计算数据范围
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const valueRange = maxValue - minValue;
  const valuePadding = valueRange * 0.1; // 10%的边距
  
  // 绘制背景
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasWidth, canvasHeight);
  
  // 绘制网格线
  ctx.strokeStyle = '#f0f0f0';
  ctx.lineWidth = 1;
  
  // 水平网格线
  for (let i = 0; i <= 5; i++) {
    const y = padding + (chartHeight / 5) * i;
    ctx.beginPath();
    ctx.moveTo(padding, y);
    ctx.lineTo(padding + chartWidth, y);
    ctx.stroke();
  }
  
  // 垂直网格线
  for (let i = 0; i <= 5; i++) {
    const x = padding + (chartWidth / 5) * i;
    ctx.beginPath();
    ctx.moveTo(x, padding);
    ctx.lineTo(x, padding + chartHeight);
    ctx.stroke();
  }
  
  // 绘制坐标轴
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, padding + chartHeight);
  ctx.lineTo(padding + chartWidth, padding + chartHeight);
  ctx.stroke();
  
  // 绘制数据点和连线
  if (values.length > 1) {
    console.log('绘制折线图:', {
      数据点数量: dates.length,
      数值范围: { minValue, maxValue, valueRange },
      图表区域: { chartWidth, chartHeight },
      坐标转换说明: 'X轴=时间位置，Y轴=数值归一化位置'
    });
    
    ctx.strokeStyle = '#5470c6';
    ctx.lineWidth = 3;
    ctx.beginPath();
    
    dates.forEach((date, index) => {
      // X轴坐标计算：时间在时间轴上的位置
      const x = padding + (chartWidth / (dates.length - 1)) * index;
      
      // Y轴坐标计算：数值在数值轴上的位置（归一化）
      const normalizedValue = (values[index] - minValue + valuePadding) / (valueRange + 2 * valuePadding);
      const y = padding + chartHeight - normalizedValue * chartHeight;
      
      console.log(`数据点${index}:`, {
        日期: date,
        数值: values[index],
        X坐标: x.toFixed(1),
        Y坐标: y.toFixed(1),
        归一化值: normalizedValue.toFixed(3)
      });
      
      if (index === 0) {
        ctx.moveTo(x, y);
      } else {
        ctx.lineTo(x, y);
      }
    });
    ctx.stroke();
    
    // 绘制数据点
    ctx.fillStyle = '#5470c6';
    dates.forEach((date, index) => {
      const x = padding + (chartWidth / (dates.length - 1)) * index;
      const normalizedValue = (values[index] - minValue + valuePadding) / (valueRange + 2 * valuePadding);
      const y = padding + chartHeight - normalizedValue * chartHeight;
      
      ctx.beginPath();
      ctx.arc(x, y, 4, 0, 2 * Math.PI);
      ctx.fill();
    });
  }
  
  // 绘制Y轴标签
  ctx.fillStyle = '#333';
  ctx.font = 'bold 14px sans-serif';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  for (let i = 0; i <= 5; i++) {
    const value = minValue + (valueRange / 5) * i;
    const y = padding + chartHeight - (chartHeight / 5) * i;
    // 确保标签位置在图表区域内且清晰可见
    if (y >= padding && y <= padding + chartHeight) {
      ctx.fillText(value.toFixed(1), padding - 15, y);
    }
  }
  
  // 绘制X轴标签
  ctx.textAlign = 'center';
  const step = Math.max(1, Math.floor(dates.length / 5));
  for (let i = 0; i < dates.length; i += step) {
    const x = padding + (chartWidth / (dates.length - 1)) * i;
    ctx.fillText(dates[i], x, padding + chartHeight + 20);
  }
  
  // 绘制标题
  ctx.fillStyle = '#333';
  ctx.font = 'bold 16px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('EMG阈值变化趋势', canvasWidth / 2, 25);
}

Page({
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
    selectedDateInfo: null, // 选中日期的详细信息
    // 图表相关（保留用于测试）
    canvasWidth: 0,
    canvasHeight: 0,
    showCanvas: false,
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

  // 初始化重试计数器
  _canvasInitRetryCount: 0,
  _maxRetryCount: 3,

  onLoad() {
    this.loadEMGData();
  },

  onShow() {
    // 页面显示时也尝试初始化canvas尺寸
    setTimeout(() => {
      if (this.data.canvasWidth === 0 || this.data.canvasHeight === 0) {
        console.log('onShow: 重新初始化canvas尺寸');
        this.initCanvasSize();
      }
    }, 200);
  },

  onReady() {
    // 延迟获取canvas尺寸，确保DOM已渲染
    setTimeout(() => {
      this.initCanvasSize();
    }, 100);
  },

  // 初始化canvas尺寸
  initCanvasSize() {
    // 防止重复初始化
    if (this._initializingCanvas) {
      console.log('Canvas正在初始化中，跳过重复调用');
      return;
    }
    
    this._initializingCanvas = true;
    
    const query = wx.createSelectorQuery();
    query.select('.chart-container').boundingClientRect((rect) => {
      console.log('图表容器尺寸查询结果:', rect);
      
      // 获取设备像素比，如果获取失败则使用默认值
      let dpr = 1;
      try {
        const deviceInfo = wx.getDeviceInfo();
        dpr = deviceInfo.pixelRatio || 1;
        console.log('设备像素比:', dpr);
      } catch (e) {
        console.log('获取设备像素比失败，使用默认值1');
        dpr = 1;
      }
      
      let canvasWidth, canvasHeight;
      
      if (rect && rect.width > 0 && rect.height > 0) {
        // 使用容器尺寸
        canvasWidth = rect.width * dpr;
        canvasHeight = rect.height * dpr;
        console.log('使用容器尺寸设置Canvas:', {
          containerWidth: rect.width,
          containerHeight: rect.height,
          dpr: dpr,
          canvasWidth: canvasWidth,
          canvasHeight: canvasHeight
        });
      } else {
        // 使用默认尺寸
        canvasWidth = 350 * dpr;
        canvasHeight = 400 * dpr;
        console.log('使用默认尺寸设置Canvas:', {
          canvasWidth: canvasWidth,
          canvasHeight: canvasHeight,
          dpr: dpr
        });
      }
      
      // 确保尺寸有效
      if (isNaN(canvasWidth) || isNaN(canvasHeight) || canvasWidth <= 0 || canvasHeight <= 0) {
        console.log('Canvas尺寸无效，使用备用尺寸');
        canvasWidth = 350;
        canvasHeight = 400;
      }
      
      this.setData({
        canvasWidth: canvasWidth,
        canvasHeight: canvasHeight
      });
      
      this._initializingCanvas = false;
      
      // 如果有数据，立即绘制图表
      if (this.data.chartDates.length > 0 && this.data.hasData) {
        this.setData({ showCanvas: true });
        setTimeout(() => {
          this.updateChart(this.data.chartDates, this.data.chartValues);
        }, 100);
      }
    }).exec();
  },

  // 加载EMG历史数据
  loadEMGData() {
    this.setData({ loading: true });
    
    // 模拟数据生成（实际项目中替换为云函数调用）
    setTimeout(() => {
      const sampleData = this.generateSampleEMGData();
      this.processEMGData(sampleData);
      this.setData({ hasData: true, loading: false });
      
      // 延迟绘制图表，确保canvas已准备好
      setTimeout(() => {
        if (this.data.chartDates.length > 0) {
          this._canvasInitRetryCount = 0; // 重置重试计数器
          this.setData({ showCanvas: true });
          this.updateChart(this.data.chartDates, this.data.chartValues);
        }
      }, 800);
    }, 1000);

    // 实际的云函数调用（暂时注释）
    /*
    wx.cloud.callFunction({
      name: 'getEMG',
      data: {
        openid: app.globalData.openId
      },
      success: (res) => {
        console.log('EMG数据获取成功:', res);
        if (res.result.success && res.result.data.length > 0) {
          this.processEMGData(res.result.data);
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
    */
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
        const threshold = (Math.random() * 40 + 30).toFixed(1); // 30-70之间的随机值
        const dateKey = this.formatDateKey(date);
        data[dateKey] = parseFloat(threshold);
      }
    }
    
    return data;
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

  // 处理EMG数据，转换为日历格式
  processEMGData(rawData) {
    // 保存所有原始数据
    this.setData({ emgData: rawData });
    this.setData({ hasData: true });
    
    // 生成日历
    this.generateCalendar();
    
    // 更新患者友好视图
    this.updatePatientFriendlyView();
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
    const weeklyValues = this.getWeeklyValues();
    const weeklyAverage = weeklyValues.length > 0 ? 
      weeklyValues.reduce((sum, val) => sum + val, 0) / weeklyValues.length : 0;
    const completion = Math.min(100, (weeklyAverage / 50) * 100);
    
    // 计算统计数据
    const totalTrainingDays = Object.keys(emgData).length;
    const bestValue = totalTrainingDays > 0 ? Math.max(...Object.values(emgData)).toFixed(1) : '0';
    
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
        average: weeklyAverage,
        target: 50,
        completion: completion
      },
      totalTrainingDays: totalTrainingDays,
      bestValue: bestValue
    });
    
    // 检查成就
    this.checkAchievements();
  },

  // 获取本周数值
  getWeeklyValues() {
    const { emgData } = this.data;
    const values = [];
    const today = new Date();
    
    for (let i = 6; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(date.getDate() - i);
      const dateKey = this.formatDateKey(date);
      if (emgData[dateKey] !== undefined) {
        values.push(emgData[dateKey]);
      }
    }
    
    return values;
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
      const maxValue = Math.max(...values);
      achievements.push({
        icon: '🌟',
        title: '最高记录',
        desc: `${maxValue.toFixed(1)}μV`
      });
    }
    
    // 检查进步最大的一天
    const dates = Object.keys(emgData).sort();
    let maxImprovement = 0;
    let bestDay = '';
    for (let i = 1; i < dates.length; i++) {
      const improvement = emgData[dates[i]] - emgData[dates[i-1]];
      if (improvement > maxImprovement) {
        maxImprovement = improvement;
        bestDay = dates[i];
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
    
    // 基于选中日期计算该周的日期范围
    const selectedDate = selectedDateInfo.date;
    const weekStart = new Date(selectedDate);
    weekStart.setDate(selectedDate.getDate() - selectedDate.getDay()); // 设置为周日
    
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6); // 设置为周六
    
    const weekStr = `${weekStart.getMonth()+1}/${weekStart.getDate()}-${weekEnd.getMonth()+1}/${weekEnd.getDate()}`;
    
    // 计算该周的所有日期
    const weekDates = [];
    for (let i = 0; i < 7; i++) {
      const date = new Date(weekStart);
      date.setDate(weekStart.getDate() + i);
      weekDates.push(this.formatDateKey(date));
    }
    
    // 更新selectedDateInfo，添加周信息
    this.setData({
      selectedDateInfo: {
        ...selectedDateInfo,
        weekStr: weekStr,
        weekStart: weekStart,
        weekEnd: weekEnd,
        weekDates: weekDates
      }
    });
    // console.log('start:',weekStart,"end:",weekEnd,"weekDates:",weekDates)
    
    // 延迟显示Toast，确保setData完成，使用简洁文本
    setTimeout(() => {
      wx.showToast({
        title: `${weekStr}`,
        icon: 'success',
        duration: 2000
      });
    }, 50);
    
    // 跳转到周图表页面
    this.navigateToWeekChart(weekStr);
  },

  // 跳转到周图表页面
  navigateToWeekChart(weekStr) {
    wx.navigateTo({
      url: `/pages/week-chart/week-chart?week=${encodeURIComponent(weekStr)}`,
      success: () => {
        wx.showToast({
          title: '跳转到周图表',
          icon: 'success',
          duration: 1500
        });
      },
      fail: () => {
        wx.showToast({
          title: '页面开发中',
          icon: 'none',
          duration: 2000
        });
      }
    });
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
    const monthStr = `${selectedDate.getFullYear()}年${selectedDate.getMonth() + 1}月`;
    
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
    }, 50);
    
    // 跳转到月图表页面
    this.navigateToMonthChart(monthStr);
  },

  // 跳转到周图表页面
  navigateToWeekChart(weekStr) {
    wx.navigateTo({
      url: `/pages/week-chart/week-chart?week=${encodeURIComponent(weekStr)}`,
      success: () => {
        wx.showToast({
          title: '跳转到周图表',
          icon: 'success',
          duration: 1500
        });
      },
      fail: () => {
        wx.showToast({
          title: '页面开发中',
          icon: 'none',
          duration: 2000
        });
      }
    });
  },

  // 跳转到月图表页面
  navigateToMonthChart(monthStr) {
    wx.navigateTo({
      url: `/pages/month-chart/month-chart?month=${encodeURIComponent(monthStr)}`,
      success: () => {
        wx.showToast({
          title: '跳转到月图表',
          icon: 'success',
          duration: 1500
        });
      },
      fail: () => {
        wx.showToast({
          title: '页面开发中',
          icon: 'none',
          duration: 2000
        });
      }
    });
  },

  // 更新当前页面的数据
  updateCurrentPageData() {
    const { emgData, currentPageIndex, pageSize } = this.data;
    
    if (!emgData || emgData.length === 0) {
      this.setData({ chartDates: [], chartValues: [] });
      return;
    }
    
    // 计算当前页的数据范围
    const startIndex = currentPageIndex * pageSize;
    const endIndex = Math.min(startIndex + pageSize, emgData.length);
    const pageData = emgData.slice(startIndex, endIndex);
    
    const dates = [];
    const values = [];
    
    pageData.forEach(item => {
      if (item.threshold && item.timestamp) {
        const date = new Date(item.timestamp);
        const dateStr = `${date.getMonth()+1}/${date.getDate()}`;
        dates.push(dateStr);
        values.push(parseFloat(item.threshold));
      }
    });

    this.setData({ 
      chartDates: dates,
      chartValues: values
    });
    
    // 更新图表
    this.updateChart(dates, values);
  },

  // 更新图表数据
  updateChart(dates, values) {
    console.log('updateChart被调用:', { 
      dates: dates ? dates.length : 0, 
      values: values ? values.length : 0,
      hasData: this.data.hasData,
      showCanvas: this.data.showCanvas
    });
    
    if (!dates || !values || dates.length === 0) {
      console.log('updateChart: 数据为空，跳过绘制');
      return;
    }
    
    // 检查canvas是否应该显示
    if (!this.data.hasData || !this.data.showCanvas) {
      console.log('Canvas不应该显示，跳过绘制');
      return;
    }
    
    // 检查canvas尺寸是否已设置
    if (this.data.canvasWidth === 0 || this.data.canvasHeight === 0 || isNaN(this.data.canvasWidth) || isNaN(this.data.canvasHeight)) {
      console.log('Canvas尺寸未设置或无效，先初始化');
      
      // 检查重试次数
      if (this._canvasInitRetryCount >= this._maxRetryCount) {
        console.log('Canvas初始化重试次数超限，使用强制默认尺寸');
        const dpr = 1; // 使用默认像素比
        const canvasWidth = 350 * dpr;
        const canvasHeight = 400 * dpr;
        
        this.setData({
          canvasWidth: canvasWidth,
          canvasHeight: canvasHeight
        });
        
        // 继续绘制
        this._drawChartDirectly(dates, values, canvasWidth, canvasHeight, dpr);
        return;
      }
      
      this._canvasInitRetryCount++;
      this.initCanvasSize();
      // 延迟重试
      setTimeout(() => {
        this.updateChart(dates, values);
      }, 500);
      return;
    }
    
    // 使用原生canvas绘制图表
    const query = wx.createSelectorQuery();
    query.select('#chart-canvas').fields({ node: true, size: true }).exec((res) => {
      console.log('canvas查询结果:', res);
      
      if (res && res[0] && res[0].node) {
        const canvas = res[0].node;
        const ctx = canvas.getContext('2d');
        
        if (!ctx) {
          console.error('无法获取canvas context');
          return;
        }
        
        // 设置canvas尺寸
        let dpr = 1;
        try {
          const deviceInfo = wx.getDeviceInfo();
          dpr = deviceInfo.pixelRatio || 1;
        } catch (e) {
          console.log('获取设备像素比失败，使用默认值1');
          dpr = 1;
        }
        
        let canvasWidth = this.data.canvasWidth;
        let canvasHeight = this.data.canvasHeight;
        
        // 如果尺寸仍然为0或NaN，强制设置默认尺寸
        if (canvasWidth === 0 || canvasHeight === 0 || isNaN(canvasWidth) || isNaN(canvasHeight)) {
          canvasWidth = 350 * dpr;
          canvasHeight = 400 * dpr;
          console.log('强制设置默认canvas尺寸:', { canvasWidth, canvasHeight, dpr });
        }
        
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        ctx.scale(dpr, dpr);
        
        // 计算显示尺寸（逻辑像素）
        const displayWidth = canvasWidth / dpr;
        const displayHeight = canvasHeight / dpr;
        
        console.log('开始绘制图表:', {
          canvasWidth: canvasWidth,
          canvasHeight: canvasHeight,
          dpr: dpr,
          dates: dates.length,
          values: values.length,
          displayWidth: displayWidth,
          displayHeight: displayHeight
        });
        
        // 绘制图表（使用显示尺寸）
        drawChart(ctx, dates, values, displayWidth, displayHeight);
      } else {
        console.error('canvas元素未找到或无效:', res);
      }
    });
  },

  // 直接绘制图表的辅助函数
  _drawChartDirectly(dates, values, canvasWidth, canvasHeight, dpr) {
    const query = wx.createSelectorQuery();
    query.select('#chart-canvas').fields({ node: true, size: true }).exec((res) => {
      console.log('直接绘制canvas查询结果:', res);
      
      if (res && res[0] && res[0].node) {
        const canvas = res[0].node;
        const ctx = canvas.getContext('2d');
        
        if (!ctx) {
          console.error('无法获取canvas context');
          return;
        }
        
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        ctx.scale(dpr, dpr);
        
        // 计算显示尺寸（逻辑像素）
        const displayWidth = canvasWidth / dpr;
        const displayHeight = canvasHeight / dpr;
        
        console.log('直接绘制图表:', {
          canvasWidth: canvasWidth,
          canvasHeight: canvasHeight,
          dpr: dpr,
          dates: dates.length,
          values: values.length,
          displayWidth: displayWidth,
          displayHeight: displayHeight
        });
        
        // 绘制图表（使用显示尺寸）
        drawChart(ctx, dates, values, displayWidth, displayHeight);
      } else {
        console.error('直接绘制时canvas元素未找到:', res);
      }
    });
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

  // 图表初始化完成回调（不再需要）
  onChartInit(e) {
    console.log('图表初始化完成');
    // 如果有数据，立即更新图表
    if (this.data.chartDates.length > 0) {
      this.updateChart(this.data.chartDates, this.data.chartValues);
    }
  },

  // 测试绘制功能（用于调试）
  testDrawChart() {
    console.log('测试绘制功能');
    const query = wx.createSelectorQuery();
    query.select('#chart-canvas').fields({ node: true, size: true }).exec((res) => {
      if (res && res[0] && res[0].node) {
        const canvas = res[0].node;
        const ctx = canvas.getContext('2d');
        
        if (!ctx) {
          console.error('无法获取canvas context');
          return;
        }
        
        // 使用固定尺寸测试
        const testWidth = 350;
        const testHeight = 400;
        canvas.width = testWidth;
        canvas.height = testHeight;
        
        console.log('测试绘制:', { testWidth, testHeight });
        
        // 测试数据
        const testDates = ['1/1', '1/2', '1/3', '1/4', '1/5'];
        const testValues = [45, 52, 48, 55, 50];
        
        // 直接绘制测试图表
        drawChart(ctx, testDates, testValues, testWidth, testHeight);
        
        wx.showToast({
          title: '测试绘制完成',
          icon: 'success'
        });
      } else {
        console.error('测试绘制时canvas元素未找到');
      }
    });
  },

  // 生成图表数据（用于图表视图）
  generateChartData() {
    const { emgData } = this.data;
    const dates = [];
    const values = [];
    
    // 获取最近30天的数据
    const now = new Date();
    for (let i = 29; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      const dateKey = this.formatDateKey(date);
      
      if (emgData[dateKey] !== undefined) {
        const dateStr = `${date.getMonth()+1}/${date.getDate()}`;
        dates.push(dateStr);
        values.push(emgData[dateKey]);
      }
    }
    
    this.setData({ chartDates: dates, chartValues: values });
    
    // 更新图表
    if (dates.length > 0) {
      setTimeout(() => {
        this._canvasInitRetryCount = 0;
        this.updateChart(dates, values);
      }, 100);
    }
  },

  // 添加测试数据点（指定某天）
  addTestDataPoint() {
    // 弹出选择器让用户选择日期
    const dates = [];
    const now = new Date();
    
    // 生成最近30天的日期选项
    for (let i = 29; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      const dateKey = this.formatDateKey(date);
      const dateStr = `${date.getMonth()+1}月${date.getDate()}日`;
      dates.push(`${dateStr} (${dateKey})`);
    }
    
    wx.showActionSheet({
      itemList: dates,
      success: (res) => {
        const selectedDateStr = dates[res.tapIndex];
        // 从字符串中提取日期key
        const dateKey = selectedDateStr.match(/\((\d{4}-\d{2}-\d{2})\)/)[1];
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
    
    // 如果当前显示图表，也更新图表数据
    if (this.data.showChart) {
      this.generateChartData();
    }
    
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
            chartDates: [],
            chartValues: [],
            hasData: false,
            showCanvas: false,
            selectedDate: null
          });
          
          // 重新生成空日历
          this.generateCalendar();
          
          wx.showToast({
            title: '数据已清空',
            icon: 'success'
          });
        }
      }
    });
  }
});

