// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()
// 云函数入口函数
exports.main = async (event, context) => {
  const {openid} = event
  try{
    const trainingData=await db.collection('training_data').where({
      openid
    }).get() 
    if(trainingData.data.length){
      return {
        event,  
        success:true,
        message:'获取用户训练数据成功',
        data:trainingData.data
      }
    }
    return {
      event,
      success:false,
      message:'该用户不存在训练数据',
      data:{} 
    }
  }catch(e){
    console.error(e)
    return {
      event,
      success:false,
      message:'获取用户训练数据失败',
      error:e
    } 
  }
}