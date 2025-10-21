// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { 
    searchKeyword = '', 
    searchType = 'all', 
    currentFilter = 'all',
    timeRange = {},
    sortBy = 'createTime',
    sortOrder = 'desc',
    currentPage = 1,
    pageSize = 10
  } = event

  try {
    // 构建查询条件
    let whereCondition = {}
    
    // 根据筛选条件添加状态过滤
    if (currentFilter === 'active') {
      whereCondition.status = 'active'
    } else if (currentFilter === 'inactive') {
      whereCondition.status = 'inactive'
    }

    // 根据搜索类型和关键词添加搜索条件
    if (searchKeyword) {
      const keyword = searchKeyword.toLowerCase()
      switch (searchType) {
        case 'openid':
          whereCondition.openid = db.RegExp({
            regexp: keyword,
            options: 'i'
          })
          break
        case 'username':
          whereCondition.username = db.RegExp({
            regexp: keyword,
            options: 'i'
          })
          break
        case 'time':
          // 时间搜索 - 这里简化处理，实际项目中需要更复杂的时间匹配
          whereCondition.createTime = db.RegExp({
            regexp: keyword,
            options: 'i'
          })
          break
        default: // 'all'
          // 使用复合查询
          whereCondition = {
            ...whereCondition,
            $or: [
              { openid: db.RegExp({ regexp: keyword, options: 'i' }) },
              { username: db.RegExp({ regexp: keyword, options: 'i' }) },
              { createTime: db.RegExp({ regexp: keyword, options: 'i' }) }
            ]
          }
          break
      }
    }

    // 时间范围过滤
    if (timeRange.start || timeRange.end) {
      if (timeRange.start) {
        whereCondition.createTime = {
          ...whereCondition.createTime,
          $gte: new Date(timeRange.start)
        }
      }
      if (timeRange.end) {
        whereCondition.createTime = {
          ...whereCondition.createTime,
          $lte: new Date(timeRange.end)
        }
      }
    }

    // 构建排序条件
    let orderBy = {}
    switch (sortBy) {
      case 'username':
        orderBy.username = sortOrder === 'asc' ? 1 : -1
        break
      case 'emgThreshold':
        orderBy.emgThreshold = sortOrder === 'asc' ? 1 : -1
        break
      case 'createTime':
      default:
        orderBy.createTime = sortOrder === 'asc' ? 1 : -1
        break
    }

    // 执行查询
    const result = await db.collection('users')
      .where(whereCondition)
      .orderBy(Object.keys(orderBy)[0], Object.values(orderBy)[0])
      .skip((currentPage - 1) * pageSize)
      .limit(pageSize)
      .get()

    // 获取总数
    const countResult = await db.collection('users')
      .where(whereCondition)
      .count()

    // 获取统计数据
    const statsResult = await db.collection('users').count()
    const activeResult = await db.collection('users')
      .where({ status: 'active' })
      .count()

    return {
      success: true,
      data: {
        userList: result.data,
        totalUsers: countResult.total,
        totalPages: Math.ceil(countResult.total / pageSize),
        userStats: {
          total: statsResult.total,
          active: activeResult.total,
          inactive: statsResult.total - activeResult.total
        }
      }
    }
  } catch (error) {
    console.error('获取用户列表失败:', error)
    return {
      success: false,
      message: '获取用户列表失败',
      error: error.message
    }
  }
}
