/*
  emg图像页面
*/

import * as echarts from '../../ec-canvas/echarts';

const app = getApp();
let chart = null;
let emgRawData = [];
let envelopData = [];
let xAxisData = [];
let dataIndex = 0;

function initChart(canvas, width, height, dpr) {
  chart = echarts.init(canvas, null, {
    width: width,
    height: height,
    devicePixelRatio: dpr 
  });
  canvas.setChart(chart);

  var option = {
    title: {
      text: 'emgRaw 与 envelop 动态曲线图',
      left: 'center'
    },
    legend: {
      data: ['emgRaw', 'envelop'],
      top: 50,
      left: 'center',
      zIndex: 100
    },
    grid: {
      containLabel: true
    },
    tooltip: {
      show: true,
      trigger: 'axis'
    },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: xAxisData
    },
    yAxis: {
      x: 'center',
      type: 'value',
      splitLine: {
        lineStyle: {
          type: 'dashed'
        }
      }
    },
    series: [{
      name: 'emgRaw',
      type: 'line',
      smooth: true,
      data: emgRawData
    }, {
      name: 'envelop',
      type: 'line',
      smooth: true,
      data: envelopData
    }]
  };

  chart.setOption(option);
  return chart;
}

function updateChart() {
  if (chart) {
    chart.setOption({
      xAxis: {
        data: xAxisData
      },
      series: [{
        name: 'emgRaw',
        data: emgRawData
      }, {
        name: 'envelop',
        data: envelopData
      }]
    });
  }
}

Page({
  onLoad() {
    this.chartComponent = this.selectComponent('#chartCanvas');
    wx.onBLECharacteristicValueChange((result) => {
      console.log('onBLECharacteristicValueChange', result.value);
      let hex = this.ab2hex(result.value);
      let dataStr = this.hextoString(hex);
      console.log('Received data:', dataStr);
      const dataArray = dataStr.split(',');
      if (dataArray.length >= 2) {
        const emgRaw = parseInt(dataArray[0]);
        const envelop = parseInt(dataArray[1]);
        emgRawData.push(emgRaw);
        envelopData.push(envelop);
        xAxisData.push(dataIndex++);
        updateChart();
      }
    });
  },
  data: {
    ec: {
      onInit: initChart
    }
  },
  ab2hex(buffer) {
    var hexArr = Array.prototype.map.call(
      new Uint8Array(buffer),
      function (bit) {
        return ('00' + bit.toString(16)).slice(-2);
      }
    );
    return hexArr.join('');
  },
  hextoString: function (hex) {
    var arr = hex.split("");
    var out = "";
    for (var i = 0; i < arr.length / 2; i++) {
      var tmp = "0x" + arr[i * 2] + arr[i * 2 + 1];
      var charValue = String.fromCharCode(parseInt(tmp, 16));
      out += charValue;
    }
    return out;
  }
});