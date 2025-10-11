// 云函数入口文件
const cloud = require('wx-server-sdk')
cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境
const db = cloud.database()
// 云函数入口函数
exports.main = async (event, context) => {
  const {openid} = event
  try{
    const user=await db.collection('user').where({
      openid
    }).get() 
    if(user.data.length){
      return {
        success:true,
        message:'获取用户信息成功',
        data:user.data[0]
      }
    }
    return {
      success:false,
      message:'用户不存在',
      data:{
        openid:openid,
        nickName:"",
        avatarURL:"",
      }
    }
  }catch(e){
    console.error(e)
    return {
      success:false,
      message:'内部服务错误',
      error:e
    }
  } 
}