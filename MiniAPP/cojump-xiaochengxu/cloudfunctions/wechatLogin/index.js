// cloudfunctions/wechatLogin/index.js
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

/**
 * 微信一键登录云函数
 * @param {object} userInfo - 微信用户信息
 */
exports.main = async (event, context) => {
  const { openId, userInfo } = event
  const wxContext = cloud.getWXContext()
  
  // 获取有效的openId
  const actualOpenId = openId ? String(openId) : String(wxContext.OPENID)
  
  // 确保openId存在
  if (!actualOpenId || actualOpenId === 'undefined') {
    console.error('无法获取openId', { openId, OPENID: wxContext.OPENID })
    return {
      success: false,
      message: '无法获取用户标识'
    }
  }
  
  // 默认名称 微信用户+openId后4位（避免暴露完整openId）
  if (!userInfo.nickName||userInfo.nickName==''||userInfo.nickName=='微信用户') {
    userInfo.nickName = '微信用户' + actualOpenId.slice(-8)
  }
  
  try {
    // 1. 查找用户是否存在
    const userQuery = await db.collection('users')
      .where({ openId: actualOpenId })
      .get()

    let user
    const currentTime = new Date().toISOString()

    if (userQuery.data.length === 0) {
      // 2. 新用户 - 创建用户记录
      const insertResult = await db.collection('users').add({
        data: {
          openId: actualOpenId,
          nickName: userInfo.nickName,
          avatarUrl: userInfo.avatarUrl || '',
          createTime: currentTime,
          lastLoginTime: currentTime,
          lastUpdateTime: currentTime
        }
      })
      
      // 获取新创建的用户信息
      const newUserQuery = await db.collection('users').doc(insertResult._id).get()
      user = newUserQuery.data
    } else {
      // 3. 老用户 - 更新登录时间和基本信息
      user = userQuery.data[0]
      
      await db.collection('users').doc(user._id).update({
        data: {
          nickName: userInfo.nickName || user.nickName,
          avatarUrl: userInfo.avatarUrl || user.avatarUrl,
          lastLoginTime: currentTime,
          lastUpdateTime: currentTime
        }
      })
      
      // 获取更新后的用户信息
      const updatedUserQuery = await db.collection('users').doc(user._id).get()
      user = updatedUserQuery.data
    }

    return {
      success: true,
      userInfo: user,
      message: '登录成功'
    }
  } catch (err) {
    console.error('微信登录失败', err)
    return {
      success: false,
      message: err.message || '登录失败'
    }
  }
}
