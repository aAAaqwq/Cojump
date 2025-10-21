// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const {username,password}=event
  console.log('登录信息：',event)
  if (username==""||password==""){
    return {
      success:false,
      message:'username或password不能为空'
    }
  }
  //检查用户是否存在
  try{
    const user=await db.collection('admin').where({
      username
    }).get()

    if(user.data.length){
      //用户存在，检验密码
      if(user.data[0].password==password){
        return {
          success:true,
          message:'登录成功'
        }
      }else{
        return {
          success:false,
          message:'密码错误'
        }
      }
     
    }else{
       //不存在
        return {
          success: false,
          message: "管理员用户账号不存在"
        }
    }
  }catch(e){
    console.error(e)
    return {
      success:false,
      message:'内部服务错误',
      error:e
    }}

}