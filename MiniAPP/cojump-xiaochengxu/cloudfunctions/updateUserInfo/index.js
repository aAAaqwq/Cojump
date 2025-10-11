// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()
// 云函数入口函数
exports.main = async (event, context) => {
  // 更新用户信息
  const { openid, avatarURL, nickName } = event
  try {
    const res=await db.collection('user').where({
      openid: openid
    }).update({
      data: {
        avatarURL,
        nickName
      }
    })
    if (res.stats.updated > 0) {
      return {
        success: true,
        message: '更新成功'
      }
    } else {
      return {
        success: false,
        message: '未找到匹配文档或数据无变化，更新失败'
      }
    }
  } catch (error) { 
    console.error('更新用户信息失败:', error)
    return {
      success: false,
      message: '内部错误更新失败'
    }
  }
}