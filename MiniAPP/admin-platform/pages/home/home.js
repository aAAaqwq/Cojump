// pages/home/home.js - 管理员用户管理页面
Page({
  data: {
    // 页面状态
    loading: false,
    pageLoading: true,

    // 搜索和筛选
    searchKeyword: '',
    currentFilter: 'all',
    sortBy: 'recordTime', // recordTime, emgThreshold, deviceId
    sortOrder: 'desc', // asc, desc

    // 高级筛选
    showAdvancedFilter: false,
    advancedFilters: {
      openid: '',
      deviceId: '',
      startTime: '',
      endTime: '',
      minEmg: '',
      maxEmg: '',
      quality: 'all'
    },

    // 数据列表
    dataList: [],// 显示的数据列表:openid,threshold,timestamp,dev_Id
    allData: [], // 存储所有数据，用于本地操作
    filteredData: [], // 存储筛选后的数据
    totalRecords: 0,
    todayRecords: 0,
    activeDevices: 0,
    abnormalRecords: 0,

    // 分页
    currentPage: 1,
    pageSize: 50,
    totalPages: 0,

    // 异常参数
    abnormalThreshold: 60,

    // 删除确认弹窗
    showDeleteModal: false,
    deleteRecordInfo: {}
  },

  onLoad() {
    this.loadData();
  },

  onShow() {
    this.getTabBar().init();
  },

  onPullDownRefresh() {
    this.loadData(true); // 强制刷新
    wx.stopPullDownRefresh();
  },

  // 搜索相关方法
  onSearchInput(e) {
    this.setData({
      searchKeyword: e.detail.value
    });
  },

  onSearch() {
    this.setData({
      currentPage: 1
    });
    this.loadData(); // 调用后端API
  },

  // 高级筛选控制
  toggleAdvancedFilter() {
    this.setData({
      showAdvancedFilter: !this.data.showAdvancedFilter
    });
  },

  hideAdvancedFilter() {
    this.setData({
      showAdvancedFilter: false
    });
  },

  // 筛选输入处理
  onFilterInput(e) {
    const field = e.currentTarget.dataset.field;
    const value = e.detail.value;
    
    this.setData({
      [`advancedFilters.${field}`]: value
    });
  },

  // 数据质量筛选
  setQuality(e) {
    const quality = e.currentTarget.dataset.quality;
    this.setData({
      'advancedFilters.quality': quality
    });
  },

  // 应用高级筛选 - 调用后端API
  applyAdvancedFilter() {
    this.setData({
      currentPage: 1,
      showAdvancedFilter: false
    });
    this.loadData(); // 调用后端API
  },

  // 重置筛选
  resetFilters() {
    this.setData({
      advancedFilters: {
        openid: '',
        deviceId: '',
        startTime: '',
        endTime: '',
        minEmg: '',
        maxEmg: '',
        quality: 'all'
      }
    });
  },

  // 排序功能 - 调用后端API
  setSort(e) {
    const field = e.currentTarget.dataset.field;
    const currentSortBy = this.data.sortBy;
    const currentSortOrder = this.data.sortOrder;
    
    if (currentSortBy === field) {
      this.setData({
        sortOrder: currentSortOrder === 'asc' ? 'desc' : 'asc'
      });
    } else {
      this.setData({
        sortBy: field,
        sortOrder: 'desc'
      });
    }
    
    this.setData({
      currentPage: 1
    });
    this.loadData(); // 调用后端API
  },

  // 筛选相关方法 - 调用后端API
  setFilter(e) {
    const filter = e.currentTarget.dataset.filter;
    this.setData({
      currentFilter: filter,
      currentPage: 1
    });
    this.loadData(); // 调用后端API
  },

  // 数据加载方法 - 使用后端数据处理
  async loadData(forceRefresh = false) {
    this.setData({ loading: true });
    console.log('searchKeyword:', this.data.searchKeyword,
      'currentFilter:', this.data.currentFilter,
      'advancedFilters:', this.data.advancedFilters,
      'sortBy:', this.data.sortBy,
      'sortOrder:', this.data.sortOrder,
      'currentPage:', this.data.currentPage,
      'pageSize:', this.data.pageSize,
      'abnormalThreshold:', this.data.abnormalThreshold);
    try {
      // 调用统一数据查询云函数
      const result = await wx.cloud.callFunction({
        name: 'getTrainingData',
        data: {
          searchKeyword: this.data.searchKeyword,
          currentFilter: this.data.currentFilter,
          advancedFilters: this.data.advancedFilters,
          sortBy: this.data.sortBy,
          sortOrder: this.data.sortOrder,
          currentPage: this.data.currentPage,
          pageSize: this.data.pageSize,
          abnormalThreshold: this.data.abnormalThreshold
        }
      });

      if (result.result.success) {
        const { dataList, pagination, statistics } = result.result.data;
        
        console.log("后端返回数据:", result.result.data);

        this.setData({
          dataList, // 直接使用后端处理后的数据
          totalRecords: statistics.totalRecords,
          todayRecords: statistics.todayRecords,
          activeDevices: statistics.activeDevices,
          abnormalRecords: statistics.abnormalRecords,
          totalPages: pagination.totalPages,
          pageLoading: false,
          loading: false
        });
      } else {
        throw new Error(result.result.message || '获取数据失败');
      }
    } catch (error) {
      console.error('加载数据失败:', error);
      this.setData({ loading: false, pageLoading: false });
      this.showToast('加载失败，请重试', 'error');
      
      // 如果云函数调用失败，使用模拟数据作为降级方案
      try {
        const mockData = await this.getMockDataList();
        this.setData({
          dataList: mockData,
          totalRecords: mockData.length,
          pageLoading: false,
          loading: false
        });
      } catch (mockError) {
        console.error('模拟数据加载也失败:', mockError);
      }
    }
  },

  // 注意：本地筛选和排序已移至后端处理，此方法不再需要
  // 所有数据处理现在通过调用loadData()方法在后端完成

  // 分页相关方法
  updatePageNumbers() {
    const { currentPage, totalPages } = this.data;
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, currentPage + 2);

    // 调整范围确保显示5个页码
    if (endPage - startPage < 4) {
      if (startPage === 1) {
        endPage = Math.min(totalPages, 5);
      } else if (endPage === totalPages) {
        startPage = Math.max(1, totalPages - 4);
      }
    }

    const pageNumbers = [];
    for (let i = startPage; i <= endPage; i++) {
      pageNumbers.push(i);
    }

    this.setData({ pageNumbers });
  },

  prevPage() {
    if (this.data.currentPage > 1) {
      this.setData({
        currentPage: this.data.currentPage - 1
      });
      this.loadData(); // 调用后端API
    }
  },

  nextPage() {
    if (this.data.currentPage < this.data.totalPages) {
      this.setData({
        currentPage: this.data.currentPage + 1
      });
      this.loadData(); // 调用后端API
    }
  },

  goToPage(e) {
    const page = e.currentTarget.dataset.page;
    this.setData({
      currentPage: page
    });
    this.loadData(); // 调用后端API
  },

  // 删除记录相关方法
  deleteRecord(e) {
    const record = e.currentTarget.dataset.item;
    console.log('删除记录:', record);
    this.setData({
      deleteRecordInfo: record,
      showDeleteModal: true
    });
  },

  hideDeleteModal() {
    this.setData({
      showDeleteModal: false,
      deleteRecordInfo: {}
    });
  },

  stopPropagation() {
    // 阻止冒泡
  },


  async confirmDelete() {
    this.setData({ loading: true });

    try {
      // 先尝试调用云函数删除记录
      const result = await wx.cloud.callFunction({
        name: 'deleteEMG',
        data: {
          _id: this.data.deleteRecordInfo._id
        }
      });

      if (result.result.success) {
        // 从本地数据中移除记录
        const allData = this.data.allData.filter(item => item._id !== this.data.deleteRecordInfo._id);
        
        this.setData({
          allData,
          showDeleteModal: false,
          deleteRecordInfo: {},
          loading: false
        });

        // 重新加载数据
        this.loadData();
        this.showToast('记录删除成功', 'success');
      } else {
        throw new Error(result.result.message || '删除记录失败');
      }
    } catch (error) {
      console.error('删除记录失败:', error);
      
      // 如果云函数调用失败，仍然从本地数据中移除（离线模式）
      const allData = this.data.allData.filter(item => item._id !== this.data.deleteRecordInfo._id);
      
      this.setData({
        allData,
        showDeleteModal: false,
        deleteRecordInfo: {},
        loading: false
      });

      // 重新加载数据
      this.loadData();
      this.showToast('记录已删除（离线模式）', 'success');
    }
  },

  // 工具方法
  getEmgClass(threshold) {
    if (!threshold) return 'emg-unknown';
    
    const value = parseFloat(threshold);
    if (value >= 80) return 'emg-excellent';
    if (value >= 60) return 'emg-good';
    return 'emg-poor';
  },

  formatTime(timeStr) {
    if (!timeStr) return '未知';

    try {
      const date = new Date(timeStr);
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    } catch (error) {
      return '未知';
    }
  },


  showToast(title, type = 'info') {
    const toast = this.selectComponent('#t-toast');
    if (toast) {
      toast.show({
        title,
        icon: type === 'success' ? 'success' : type === 'error' ? 'error' : 'info'
      });
    } else {
      wx.showToast({
        title,
        icon: type === 'success' ? 'success' : type === 'error' ? 'none' : 'none'
      });
    }
  },
  setMockData() {
    const mockData = this.getMockDataList();
    for (let i = 0; i < mockData.length; i++) {
      wx.cloud.callFunction({
        name: 'setEMG',
        data: {
          openid: mockData[i].openid,
          emgRaw: mockData[i].emgRaw,
          emgThreshold: mockData[i].emgThreshold,
          recordTime: mockData[i].recordTime,
          deviceId: mockData[i].deviceId
        }
      }).then(res => {
        console.log('设置模拟数据成功:', res);
      }).catch(err => {
        console.error('设置模拟数据失败:', err);
      });
    }
  },
  getMockDataList() {
    let mockData = [];
    // 模拟3个患者近一个月的康复数据
    for (let i = 0; i < 3; i++) {
      for (let j = 0; j < 30; j++) {
        let openId = 'openId_' + i;
        if (i === 0) {
          openId = 'oIv9712dd4TVL6MftWKmwmcvxiQ4';
        }
        mockData.push({
          openid: openId,
          emgRaw: Math.random() * 100,
          emgThreshold: (Math.random() * 30 + 40).toFixed(1),
          recordTime: this.formatTime(new Date().setDate(new Date().getDate() - j)).toISOString(),
          deviceId: 'DEV00' + i
        });
      }
    }
    return mockData;
  }
})