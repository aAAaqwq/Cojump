// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { userId, updateData } = event

  if (!userId) {
    return {
      success: false,
      message: '用户ID不能为空'
    }
  }

  try {
    // 更新用户信息
    const result = await db.collection('users')
      .doc(userId)
      .update({
        data: {
          ...updateData,
          updateTime: new Date()
        }
      })

    if (result.stats.updated > 0) {
      return {
        success: true,
        message: '用户信息更新成功',
        data: result
      }
    } else {
      return {
        success: false,
        message: '用户不存在或更新失败'
      }
    }
  } catch (error) {
    console.error('更新用户信息失败:', error)
    return {
      success: false,
      message: '更新用户信息失败',
      error: error.message
    }
  }
}
