// 云函数入口文件
const cloud = require('wx-server-sdk')
const crypto = require('crypto')

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV }) // 使用当前云环境

const db = cloud.database()

// 防暴力破解配置
const USER_COOLDOWN = 3 * 1000; // 用户登录冷却时间：3秒
const MAX_USER_ATTEMPTS = 8; // 用户最大尝试次数：8次
const USER_LOCKOUT_DURATION = 5 * 60 * 1000; // 用户锁定时间：5分钟

// 密码哈希函数
function hashPassword(password, salt) {
  return crypto.pbkdf2Sync(password, salt, 10000, 64, 'sha512').toString('hex');
}

// 生成盐值
function generateSalt() {
  return crypto.randomBytes(32).toString('hex');
}

// 验证密码
function verifyPassword(password, hashedPassword, salt) {
  const hash = hashPassword(password, salt);
  return hash === hashedPassword;
}

// 检查用户登录频率限制（基于OPENID）
async function checkUserRateLimit(openId) {
  const now = Date.now();
  const userKey = `login_limit_${openId}`;

  try {
    // 使用 where().get() 查询记录，不存在时返回空数组
    const userData = await db.collection('login_rate_limit').where({
      _id: userKey
    }).get();

    // 如果记录不存在，创建新记录
    if (!userData.data || userData.data.length === 0) {
      console.log('用户首次登录，创建限制记录');
      await db.collection('login_rate_limit').add({
        data: {
          _id: userKey,
          openId: openId,
          lastAttempt: now,
          attemptCount: 1,
          lockTime: 0
        }
      });
      return { allowed: true };
    }

    // 记录存在，获取数据
    const record = userData.data[0];
    const lastAttempt = record.lastAttempt || 0;
    const attemptCount = record.attemptCount || 0;
    const lockTime = record.lockTime || 0;

    // 检查是否在用户锁定期内
    if (lockTime > 0 && now - lockTime < USER_LOCKOUT_DURATION) {
      const remainingTime = Math.ceil((USER_LOCKOUT_DURATION - (now - lockTime)) / 60000);
      return {
        allowed: false,
        message: `账户已被锁定，请${remainingTime}分钟后再试`
      };
    }

    // 检查登录频率（3秒冷却）
    if (now - lastAttempt < USER_COOLDOWN) {
      return {
        allowed: false,
        message: '登录过于频繁，请3秒后再试'
      };
    }

    // 更新尝试次数
    const newAttemptCount = attemptCount + 1;
    const shouldLock = newAttemptCount >= MAX_USER_ATTEMPTS;

    await db.collection('login_rate_limit').doc(record._id).update({
      data: {
        lastAttempt: now,
        attemptCount: shouldLock ? 0 : newAttemptCount,
        lockTime: shouldLock ? now : 0
      }
    });

    if (shouldLock) {
      return {
        allowed: false,
        message: `连续${MAX_USER_ATTEMPTS}次登录失败，账户已被锁定5分钟`
      };
    }
    
    return { allowed: true };
  } catch (error) {
    console.error('检查用户频率限制失败:', error);
    // 数据库错误时，为了安全考虑，拒绝登录
    return { 
      allowed: false, 
      message: '系统繁忙，请稍后重试' 
    };
  }
}

// 清除用户成功登录后的记录
async function clearUserSuccess(openId) {
  try {
    await db.collection('login_rate_limit').doc(`login_limit_${openId}`).remove();
  } catch (error) {
    console.error('清除用户成功记录失败:', error);
  }
}

// 更新最后登录时间，更新openId
async function updateLastLoginInfo(username, openId) {
  try {
    const result = await db.collection('admin').doc(username).update({
      data: {
        lastLoginTime: Date.now(),
        openId: openId,
      }
    });
    return result.stats.updated > 0;
  } catch (error) {
    console.error('更新最后登录信息:', error);
    return false;
  }
}


// 检查数据格式
function checkDataFormat(username, password) {
  if (!username || !password) {
    return false;
  }
  if (username.length < 3 || password.length < 6) {
    return false;
  }
  return true;
}

// 云函数入口函数
exports.main = async (event, context) => {
  const { username, password } = event;
  const wxContext = cloud.getWXContext()
  const openId = wxContext.OPENID;

  if (!openId) {
    return {
      success: false,
      message: 'openId不能为空'
    };
  }

  console.log('登录信息：', { username, openId });

  // 基本验证
  if (!username || !password) {
    return {
      success: false,
      message: '用户名或密码不能为空'
    };
  }

  if (!checkDataFormat(username, password)) {
    return {
      success: false,
      message: '用户名或密码格式不正确'
    };
  }
  try {
    // 检查用户登录频率限制（基于OPENID）

    const loginLimitCheck = await checkUserRateLimit(openId);
    if (!loginLimitCheck.allowed) {
      return {
        success: false,
        message: loginLimitCheck.message
      };
    }

    // 查询用户
    const user = await db.collection('admin').where({
      username: username
    }).get();

    if (user.data.length === 0) {
      return {
        success: false,
        message: '管理员用户账号不存在'
      };
    }

    const userData = user.data[0];

    // 验证密码（支持明文和哈希密码）
    let passwordValid = false;

    if (userData.passwordHash && userData.salt) {
      // 使用哈希密码验证
      passwordValid = verifyPassword(password, userData.passwordHash, userData.salt);
    }

    if (passwordValid) {
      // 登录成功，清除用户限制记录
      if (openId) {
        await clearUserSuccess(openId);
      }

      await updateLastLoginInfo(username, openId);

      return {
        success: true,
        message: '登录成功',
        userInfo: {
          username: userData.username,
          role: userData.role || 'admin',
          lastLogin: Date.now()
        }
      };
    } else {
      // 登录失败，返回错误信息
      return {
        success: false,
        message: '账号或密码错误'
      };
    }

  } catch (error) {
    console.error('登录验证失败:', error);
    return {
      success: false,
      message: '内部服务错误，请稍后重试'
    };
  }
};