// cloudfunctions/updateUserInfo/index.js
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

/**
 * 更新用户信息云函数
 * @param {string} openId - 用户openId
 * @param {string} nickName - 新昵称（可选）
 * @param {string} avatarUrl - 新头像URL（可选，云存储地址）
 */
exports.main = async (event, context) => {
  const { openId, nickName, avatarUrl } = event
  const wxContext = cloud.getWXContext()
  const actualOpenId = openId || wxContext.OPENID
  
  try {
    // 1. 验证参数
    if (!actualOpenId) {
      return {
        success: false,
        message: '缺少openId参数'
      }
    }

    // 2. 查找用户
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
    const userId = user._id

    // 3. 准备更新数据
    const updateData = {}

    // 检查并处理昵称：去除首尾空白，确保不是空字符串
    if (nickName && nickName.trim() !== '') {
      updateData.nickName = nickName.trim()
    }

    if (avatarUrl) {
      updateData.avatarUrl = avatarUrl
      
      // 删除旧头像（如果存在且是云存储地址）
      if (user.avatarUrl && user.avatarUrl.startsWith('cloud://')) {
        try {
          await cloud.deleteFile({
            fileList: [user.avatarUrl]
          })
          console.log('旧头像已删除:', user.avatarUrl)
        } catch (deleteErr) {
          console.error('删除旧头像失败:', deleteErr)
          // 不影响主流程
        }
      }
    }

    // 4. 更新数据库
    await db.collection('users').doc(userId).update({
      data: updateData
    })

    // 5. 获取更新后的用户信息
    const updatedUserQuery = await db.collection('users').doc(userId).get()
    const updatedUserInfo = updatedUserQuery.data

    return {
      success: true,
      message: '更新成功',
      userInfo: updatedUserInfo
    }
  } catch (err) {
    console.error('更新用户信息失败', err)
    return {
      success: false,
      message: err.message || '更新失败'
    }
  }
}
