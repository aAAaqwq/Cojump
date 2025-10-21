// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

function formatToOneDecimalPlace(input) {
  // 处理 null 和 undefined
  if (input === null || input === undefined) {
    return '0.0'
  }
  
  // 处理空字符串
  if (input === '') {
    return '0.0'
  }
  
  // 转换为数字
  const number = parseFloat(input)
  
  // 检查是否为有效数字
  if (isNaN(number)) {
    console.warn('无效的数字输入:', input)
    return '0.0'
  }
  
  // 检查是否为无穷大
  if (!isFinite(number)) {
    console.warn('无穷大的数字输入:', input)
    return '0.0'
  }
  
  // 保留一位小数并返回字符串
  return number.toFixed(1)
}
// 云函数入口函数
exports.main = async (event, context) => {
  const { openId, deviceId, emgRaw, emgThreshold, recordTime} = event
  // 保留1位小数，确保类型为字符串
  const emgThresholdFixed = formatToOneDecimalPlace(emgThreshold);
  try {
    // 直接添加新的训练数据记录
    const res = await db.collection('training_data').add({
      data: {
        openId,
        deviceId,
        emgRaw,
        emgThreshold: emgThresholdFixed,
        recordTime
      }
    })
    if (res) {
      return {
        success: true,
        message: '添加训练数据成功',
        data: res
      }
    } else {
      return {
        success: false,
        message: '添加训练数据失败',
        error: '返回的ID为空'
      }
    }
  } catch (e) {
    console.error(e)
    return {
      success: false,
      message: '内部错误,添加训练数据失败',
      error: e
    }
  }
}