// cloudfunctions/getUserInfo/index.js
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

/**
 * 获取用户信息云函数
 * @param {string} openId - 用户openId
 */
exports.main = async (event, context) => {
  const { openId } = event
  const wxContext = cloud.getWXContext()
  const actualOpenId = openId || wxContext.OPENID

  if (!actualOpenId) {
    return {
      success: false,
      message: '无法获取用户OpenID'
    }
  }

  try {
    // 查询用户信息
    const userQuery = await db.collection('users')
      .where({ openId: actualOpenId })
      .get()

    if (userQuery.data.length === 0) {
      return {
        success: false,
        message: '用户不存在'
      }
    }

    const user = userQuery.data[0]

    return {
      success: true,
      userInfo: user,
      message: '获取成功'
    }
  } catch (e) {
    console.error('getUserInfo cloud function error', e)
    return {
      success: false,
      message: `获取失败: ${e.message}`
    }
  }
}