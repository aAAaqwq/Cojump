// 云函数：getTrainingData - 统一数据查询接口
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV })
const db = cloud.database()

const TableName = 'training_data'

exports.main = async (event, context) => {
  const {
    // 筛选参数
    searchKeyword = '',
    currentFilter = 'all', // all, today, week, month, abnormal
    advancedFilters = {},
    
    // 排序参数
    sortBy = 'recordTime', // recordTime, emgThreshold, deviceId
    sortOrder = 'desc', // asc, desc
    
    // 分页参数
    currentPage = 1,
    pageSize = 50,

    // 异常参数
    abnormalThreshold = 60
  } = event



  //To Do: 接口鉴权

  try {
    // 1. 构建查询条件
    const whereCondition = buildWhereCondition(currentFilter, searchKeyword, advancedFilters,abnormalThreshold)
    
    console.log('whereCondition:', whereCondition);
    console.log('sortBy:', sortBy, 'sortOrder:', sortOrder);
    
    // 3. 执行数据查询
    const dataResult = await db.collection(TableName)
      .where(whereCondition)
      .orderBy(sortBy, sortOrder)
      .skip((currentPage - 1) * pageSize)
      .limit(pageSize)
      .get()
    
    // 4. 计算统计数据
    const statistics = await calculateStatistics(abnormalThreshold)
    
    console.log('统计数据:', statistics);
    
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
    const todayStr = today.toISOString().split('T')[0] // 2025-10-21
    const tomorrow = new Date(today)
    tomorrow.setDate(tomorrow.getDate() + 1)
    const tomorrowStr = tomorrow.toISOString().split('T')[0] // 2025-10-22
    
    whereCondition.recordTime = db.command.and([
      db.command.gte(todayStr),     // >= "2025-10-21"
      db.command.lt(tomorrowStr)    // < "2025-10-22"
    ])
  } else if (currentFilter === 'week') {
    const weekAgo = new Date()
    weekAgo.setDate(weekAgo.getDate() - 7)
    const weekAgoStr = weekAgo.toISOString().split('T')[0] // 2025-10-14
    whereCondition.recordTime = db.command.gte(weekAgoStr)
  } else if (currentFilter === 'month') {
    // 获取本月第一天
    const now = new Date()
    const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1)
    const firstDayStr = firstDayOfMonth.toISOString().split('T')[0] // 2025-10-01
    
    whereCondition.recordTime = db.command.gte(firstDayStr)
  } else if (currentFilter === 'abnormal') {
    whereCondition.emgThreshold = db.command.lt(abnormalThreshold)
  }
  
  // 搜索关键词
  if (searchKeyword) {
    whereCondition = db.command.or([
      { openId: db.RegExp({ regexp: searchKeyword, options: 'i' }) }, //正则表达式匹配：不区分大小写
      { deviceId: db.RegExp({ regexp: searchKeyword, options: 'i' }) }
    ])
  }
  
  // 高级筛选
  if (advancedFilters.openId) {
    whereCondition.openId = db.RegExp({ 
      regexp: advancedFilters.openId, 
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
    const timeConditions = []
    if (advancedFilters.startTime) {
      // 将日期转换为字符串格式进行比较
      const startDate = new Date(advancedFilters.startTime)
      const startDateStr = startDate.toISOString().split('T')[0]
      timeConditions.push(db.command.gte(startDateStr))
    }
    if (advancedFilters.endTime) {
      // 将日期转换为字符串格式进行比较
      const endDate = new Date(advancedFilters.endTime)
      const endDateStr = endDate.toISOString().split('T')[0]
      timeConditions.push(db.command.lte(endDateStr))
    }
    whereCondition.recordTime = db.command.and(timeConditions)
  }
  
  if (advancedFilters.minEmg || advancedFilters.maxEmg) {
    const emgConditions = []
    if (advancedFilters.minEmg) {
      emgConditions.push(db.command.gte(parseFloat(advancedFilters.minEmg)))
    }
    if (advancedFilters.maxEmg) {
      emgConditions.push(db.command.lte(parseFloat(advancedFilters.maxEmg)))
    }
    whereCondition.emgThreshold = db.command.and(emgConditions)
  }
  
  return whereCondition
}

// 构建排序条件
function buildOrderBy(sortBy, sortOrder) {
  const orderBy = {}
  
  switch (sortBy) {
    case 'emgThreshold':
      orderBy.emgThreshold = sortOrder
      break
    case 'deviceId':
      orderBy.deviceId = sortOrder
      break
    case 'recordTime':
    default:
      orderBy.recordTime = sortOrder
      break
  }
  
  return orderBy
}

// 计算统计数据
async function calculateStatistics(abnormalThreshold) {
  try {
    const [totalResult, todayResult, abnormalResult, deviceResult] = await Promise.all([
      // 总记录数
      db.collection(TableName).count(),
      
      // 今日记录数
      db.collection(TableName).where({
        recordTime: db.command.and([
          db.command.gte(new Date().toISOString().split('T')[0]),
          db.command.lt(new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().split('T')[0])
        ])
      }).count(),
      
      // 异常记录数
      db.collection(TableName).where({
        emgThreshold: db.command.lt(abnormalThreshold)
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
