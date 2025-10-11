// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { openid, dev_id, emg_raw, timestamp,threshold} = event

  try {
    // 直接添加新的训练数据记录
    const res = await db.collection('training_data').add({
      data: {
        openid,
        dev_id,
        emg_raw,
        threshold,
        timestamp
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