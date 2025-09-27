// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const {openid,nickName,avatarURL}=event
  console.log('登录信息：',event)
  if (openid==""){
    return {
      success:false,
      message:'openid不能为空'
    }
  }
  //检查用户是否存在
  try{
    const user=await db.collection('user').where({
      openid
    }).get()
    if(user.data.length){
      //用户存在则只更新用户信息
      await db.collection('user').where({
        openid
      }).update({
        data:{
          nickName,
          avatarURL
        }
      })
    }else{
       //不存在则添加
      await db.collection('user').add({
        data:{
        openid:openid,
        nickName:nickName,
        avatarURL:avatarURL,
        }
    })
    }
  }catch(e){
    console.error(e)
    return {
      success:false,
      message:'登录失败',
      error:e
    }
  }
  return {
    success:  true,
    message:"登录成功",
  }
}