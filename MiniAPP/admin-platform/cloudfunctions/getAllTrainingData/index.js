// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()
// 云函数入口函数
exports.main = async (event, context) => {
  // 防止非法访问接口



  // 业务处理

  try{
    const trainingData=await db.collection('training_data').get()
    if(trainingData.data.length){
      return {
        success:true,
        message:'获取训练数据成功',
        data:trainingData.data
      }
    }
    return {
      success:false,
      message:'获取训练数据失败',
      data:{}
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