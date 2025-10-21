// 云函数：getTrainingData - 统一数据查询接口
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV })
const db = cloud.database()

exports.main = async (event, context) => {
  const {
    // 筛选参数
    searchKeyword = '',
    currentFilter = 'all', // all, today, week, month, abnormal
    advancedFilters = {},
    
    // 排序参数
    sortBy = 'recordTime', // recordTime, emgThreshold, deviceId
    sortOrder = 'asc', // asc, desc
    
    // 分页参数
    currentPage = 1,
    pageSize = 50,

    // 异常参数
    abnormalThreshold = 60
  } = event

  const TableName = 'training_data'

  //To Do: 接口鉴权

  try {
    // 1. 构建查询条件
    const whereCondition = buildWhereCondition(currentFilter, searchKeyword, advancedFilters)
    
    // 2. 构建排序条件
    const orderBy = buildOrderBy(sortBy, sortOrder)

    console.log('whereCondition:', whereCondition);
    console.log('orderBy:', orderBy);
    
    // 3. 执行数据查询
    const dataResult = await db.collection(TableName)
      .where(whereCondition)
      .orderBy(Object.keys(orderBy)[0], Object.values(orderBy)[0])
      .skip((currentPage - 1) * pageSize)
      .limit(pageSize)
      .get()
    
    // 4. 计算统计数据
    const statistics = await calculateStatistics()
    
    return {
      success: true,
      data: {
        // 数据列表
        dataList: dataResult.data,
        
        // 分页信息
        pagination: {
          currentPage,
          pageSize,
          totalRecords: statistics.totalRecords,
          totalPages: Math.ceil(statistics.totalRecords / pageSize)
        },
        
        // 统计信息
        statistics: {
          totalRecords: statistics.totalRecords,
          todayRecords: statistics.todayRecords,
          activeDevices: statistics.activeDevices,
          abnormalRecords: statistics.abnormalRecords
        }
      }
    }
  } catch (error) {
    console.error('获取数据失败:', error)
    return { 
      success: false, 
      message: '获取数据失败',
      error: error.message 
    }
  }
}

// 构建查询条件
function buildWhereCondition(currentFilter, searchKeyword, advancedFilters, abnormalThreshold) {
  let whereCondition = {}
  
  // 基础筛选
  if (currentFilter === 'today') {
    const today = new Date()
    today.setHours(0, 0, 0, 0)
    const tomorrow = new Date(today)
    tomorrow.setDate(tomorrow.getDate() + 1)
    
    whereCondition.recordTime = {
      $gte: today,   // >=今天
      $lt: tomorrow  // <明天
    }
  } else if (currentFilter === 'week') {
    const weekAgo = new Date()
    weekAgo.setDate(weekAgo.getDate() - 7)
    whereCondition.recordTime = { $gte: weekAgo }
  } else if (currentFilter === 'month') {
    const monthAgo = new Date()
    monthAgo.setMonth(monthAgo.getMonth() - 1)
    whereCondition.recordTime = { $gte: monthAgo }
  } else if (currentFilter === 'abnormal') {
    whereCondition.emgThreshold = { $lt: abnormalThreshold }
  }
  
  // 搜索关键词
  if (searchKeyword) {
    whereCondition.$or = [
      { openid: db.RegExp({ regexp: searchKeyword, options: 'i' }) }, //正则表达式匹配：不区分大小写
      { deviceId: db.RegExp({ regexp: searchKeyword, options: 'i' }) }
    ]
  }
  
  // 高级筛选
  if (advancedFilters.openid) {
    whereCondition.openid = db.RegExp({ 
      regexp: advancedFilters.openid, 
      options: 'i' 
    })
  }
  
  if (advancedFilters.deviceId) {
    whereCondition.deviceId = db.RegExp({ 
      regexp: advancedFilters.deviceId, 
      options: 'i' 
    })
  }
  
  if (advancedFilters.startTime || advancedFilters.endTime) {
    whereCondition.recordTime = {}
    if (advancedFilters.startTime) {
      whereCondition.recordTime.$gte = new Date(advancedFilters.startTime)
    }
    if (advancedFilters.endTime) {
      whereCondition.recordTime.$lte = new Date(advancedFilters.endTime)
    }
  }
  
  if (advancedFilters.minEmg || advancedFilters.maxEmg) {
    whereCondition.emgThreshold = {}
    if (advancedFilters.minEmg) {
      whereCondition.emgThreshold.$gte = parseFloat(advancedFilters.minEmg)
    }
    if (advancedFilters.maxEmg) {
      whereCondition.emgThreshold.$lte = parseFloat(advancedFilters.maxEmg)
    }
  }
  
  return whereCondition
}

// 构建排序条件
function buildOrderBy(sortBy, sortOrder) {
  const orderBy = {}
  const direction = sortOrder === 'asc' ? 1 : -1
  
  switch (sortBy) {
    case 'emgThreshold':
      orderBy.emgThreshold = direction
      break
    case 'deviceId':
      orderBy.deviceId = direction
      break
    case 'recordTime':
    default:
      orderBy.recordTime = direction
      break
  }
  
  return orderBy
}

// 计算统计数据
async function calculateStatistics() {
  try {
    const [totalResult, todayResult, abnormalResult, deviceResult] = await Promise.all([
      // 总记录数
      db.collection(TableName).count(),
      
      // 今日记录数
      db.collection(TableName).where({
        recordTime: {
          $gte: new Date(new Date().setHours(0, 0, 0, 0)),
          $lt: new Date(new Date().setHours(23, 59, 59, 999))
        }
      }).count(),
      
      // 异常记录数
      db.collection(TableName).where({
        emgThreshold: { $lt: abnormalThreshold }
      }).count(),
      
      // 活跃设备数
      db.collection(TableName).aggregate()
        .group({ _id: '$deviceId' })
        .end()
    ])
    
    return {
      totalRecords: totalResult.total,
      todayRecords: todayResult.total,
      activeDevices: deviceResult.list.length,
      abnormalRecords: abnormalResult.total
    }
  } catch (error) {
    console.error('计算统计数据失败:', error)
    return {
      totalRecords: 0,
      todayRecords: 0,
      activeDevices: 0,
      abnormalRecords: 0
    }
  }
}
