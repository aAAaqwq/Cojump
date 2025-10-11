import * as echarts from '../../ec-canvas/echarts';

const app = getApp();

function initChart(canvas, width, height, dpr) {
  const chart = echarts.init(canvas, null, {
    width: width,
    height: height,
    devicePixelRatio: dpr
  });
  canvas.setChart(chart);

  // 初始化为空图表，等待数据加载
  var option = {
    title: {
      text: 'EMG阈值历史趋势',
      left: 'center',
      textStyle: {
        fontSize: 16,
        fontWeight: 'bold'
      }
    },
    grid: {
      containLabel: true,
      left: '10%',
      right: '10%',
      top: '15%',
      bottom: '15%'
    },
    tooltip: {
      show: true,
      trigger: 'axis',
      formatter: function(params) {
        if (params && params.length > 0) {
          const data = params[0];
          const date = new Date(data.value[0]);
          const timeStr = date.toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
          });
          return `${timeStr}<br/>EMG阈值: ${data.value[1]}`;
        }
        return '';
      }
    },
    xAxis: {
      type: 'time',
      boundaryGap: false,
      axisLabel: {
        formatter: function(value) {
          const date = new Date(value);
          return `${date.getMonth()+1}/${date.getDate()}`;
        }
      }
    },
    yAxis: {
      type: 'value',
      name: '阈值',
      nameLocation: 'middle',
      nameGap: 30,
      splitLine: {
        lineStyle: {
          type: 'dashed',
          color: '#f0f0f0'
        }
      }
    },
    series: [{
      name: 'EMG阈值',
      type: 'line',
      smooth: true,
      symbol: 'circle',
      symbolSize: 4,
      lineStyle: {
        color: '#5470c6',
        width: 2
      },
      itemStyle: {
        color: '#5470c6'
      },
      areaStyle: {
        color: {
          type: 'linear',
          x: 0,
          y: 0,
          x2: 0,
          y2: 1,
          colorStops: [{
            offset: 0, color: 'rgba(84, 112, 198, 0.3)'
          }, {
            offset: 1, color: 'rgba(84, 112, 198, 0.05)'
          }]
        }
      },
      data: []
    }]
  };

  chart.setOption(option);
  return chart;
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
    ec: {
      onInit: initChart
    },
    chart: null,
    timeRange: 'week', // 默认显示一周数据
    loading: false,
    hasData: false,
    emgData: []
  },

  onLoad() {
    this.loadEMGData();
  },

  onReady() {
  },

  // 加载EMG历史数据
  loadEMGData() {
    this.setData({ loading: true });
    
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
  },

  // 处理EMG数据，转换为图表格式
  processEMGData(rawData) {
    const chartData = rawData
      .filter(item => item.threshold && item.timestamp) // 过滤有效数据
      .map(item => [
        new Date(item.timestamp).getTime(), // x轴：时间戳
        parseFloat(item.threshold) // y轴：阈值
      ])
      .sort((a, b) => a[0] - b[0]); // 按时间排序

    this.setData({ emgData: chartData });
    this.updateChart(chartData);
  },

  // 更新图表数据
  updateChart(data) {
    if (!this.data.chart) return;
    
    this.data.chart.setOption({
      series: [{
        data: data
      }]
    });
  },

  // 时间范围切换
  switchTimeRange(e) {
    const range = e.currentTarget.dataset.range;
    this.setData({ timeRange: range });
    
    // 根据时间范围筛选数据
    let filteredData = this.data.emgData;
    const now = new Date().getTime();
    
    switch(range) {
      case 'day':
        const dayAgo = now - 24 * 60 * 60 * 1000;
        filteredData = this.data.emgData.filter(item => item[0] >= dayAgo);
        break;
      case 'week':
        const weekAgo = now - 7 * 24 * 60 * 60 * 1000;
        filteredData = this.data.emgData.filter(item => item[0] >= weekAgo);
        break;
      case 'month':
        const monthAgo = now - 30 * 24 * 60 * 60 * 1000;
        filteredData = this.data.emgData.filter(item => item[0] >= monthAgo);
        break;
      default:
        filteredData = this.data.emgData;
    }
    
    this.updateChart(filteredData);
  },

  // 刷新数据
  refreshData() {
    this.loadEMGData();
  },

  // 图表初始化完成回调
  onChartInit(e) {
    this.setData({ chart: e.detail.chart });
    // 如果有数据，立即更新图表
    if (this.data.emgData.length > 0) {
      this.updateChart(this.data.emgData);
    }
  }
});
