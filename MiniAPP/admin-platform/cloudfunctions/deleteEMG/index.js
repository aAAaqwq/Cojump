// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { _id}=event
  if (!_id) {
    return {
      success: false,
      message: 'id不能为空'
    }
  }

  try {
    const result = await db.collection('training_data').where({ _id }).remove()
    if (result.stats.removed > 0) {
      return {
        success: true,
        message: '删除训练数据成功',
        data: result
      }
    } else {
      return {
        success: false,
        message: '删除训练数据失败'
      }
    }
  } catch (error) {
    console.error('删除训练数据失败:', error)
    return {
      success: false,
      message: '服务器内部错误',
      error: error.message
    }
  }
  
}