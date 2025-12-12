// 月度销售分析页面
export class MonthlyPage {
    constructor(app) {
        this.app = app;
        this.currentFilePath = null;
        this.fileOptions = null;
        this.analysisResult = null;
        this.isDataLoaded = false;
        this.startTime = null;
        this.timerInterval = null;
        this.unlistenProgress = null;
        this.currentDimension = 'customer';
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">📈</span>
                        月度销售分析
                    </h1>
                    <p class="page-desc">
                        使用已导入的数据源，按客户或地区分析月度销售趋势
                    </p>
                </div>
                
                <div class="data-source-notice slide-up" id="dataSourceNotice" style="display: none;">
                    <div class="notice-card">
                        <div class="notice-icon">⚠️</div>
                        <div class="notice-content">
                            <h4>未导入数据源</h4>
                            <p>请先在首页导入数据源，然后才能进行分析</p>
                            <button class="btn btn-primary" id="goToHomeBtn">前往首页导入</button>
                        </div>
                    </div>
                </div>
                
                <div class="upload-section slide-up" id="uploadSection" style="display: none;">
                    <div class="data-source-info-card">
                        <div class="ds-info-header">
                            <div class="ds-select-group">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                    <label class="ds-info-label">选择数据源（可多选合并分析）：</label>
                                    <button type="button" class="btn-select-all" id="selectAllBtn" style="padding: 4px 12px; font-size: 12px; background: var(--accent-blue); color: white; border: none; border-radius: 4px; cursor: pointer;">
                                        全选
                                    </button>
                                </div>
                                <div class="data-source-checkboxes" id="dataSourceCheckboxes">
                                    <p style="color: var(--text-muted);">加载中...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="info-box">
                        <h4>📌 所需列名说明</h4>
                        <ul>
                            <li><strong>客户编码</strong> - 用于识别唯一客户（必需）</li>
                            <li><strong>支付金额</strong> - 支付金额数值（必需）</li>
                            <li><strong>充值抵扣</strong> - 充值抵扣金额（必需）</li>
                            <li><strong>日期/下单日期/出库时间</strong> - 用于按月汇总（必需，支持：日期、订单日期、下单日期、出库时间、出库日期、发货时间等）</li>
                            <li><strong>省/市/区</strong> - 地区信息（可选）</li>
                        </ul>
                    </div>
                </div>
                
                <!-- 分析选项 -->
                <div class="analysis-options" id="analysisOptions" style="display: none;">
                    <div class="options-card slide-up">
                        <div class="options-header">
                            <h3>📊 选择分析条件</h3>
                            <span class="cache-status" id="cacheStatus">✅ 数据已缓存</span>
                        </div>
                        
                        <!-- 分析维度 -->
                        <div class="filter-row">
                            <div class="filter-item">
                                <label>📊 分析维度：</label>
                                <div class="option-tabs">
                                    <button class="option-tab active" data-type="customer">按客户</button>
                                    <button class="option-tab" data-type="province">按省份</button>
                                    <button class="option-tab" data-type="city">按城市</button>
                                    <button class="option-tab" data-type="district">按区县</button>
                                    <button class="option-tab" data-type="region">按完整地区</button>
                                </div>
                            </div>
                        </div>
                        
                        <!-- 筛选目标 -->
                        <div class="filter-row" id="targetFilterRow">
                            <div class="filter-item">
                                <label id="targetLabel">🎯 选择客户：</label>
                                <div class="target-input-wrapper">
                                    <input type="text" id="targetInput" class="select-input" placeholder="输入关键词搜索或选择..." autocomplete="off">
                                    <div class="target-dropdown" id="targetDropdown" style="display: none;"></div>
                                </div>
                            </div>
                        </div>
                        
                        <button class="btn btn-primary" id="analyzeBtn" disabled>
                            <span>🔍</span>
                            开始分析
                        </button>
                    </div>
                </div>
                
                <!-- 结果区域 -->
                <div class="result-section" id="resultSection">
                    <div class="result-header">
                        <div>
                            <h2 class="result-title">
                                <span>📈</span>
                                <span id="resultTitle">月度销售趋势</span>
                            </h2>
                            <p class="result-subtitle" id="resultSubtitle"></p>
                        </div>
                        <div class="stats-row">
                            <div class="stat-card">
                                <div class="stat-label">总金额</div>
                                <div class="stat-value" id="totalAmount">¥0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">总订单数</div>
                                <div class="stat-value" id="totalOrders">0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">月份数</div>
                                <div class="stat-value" id="monthCount">0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">分析耗时</div>
                                <div class="stat-value" id="processTime">0ms</div>
                            </div>
                        </div>
                        <div style="display: flex; gap: 12px;">
                            <button class="btn btn-primary" id="exportBtn">
                                <span>📥</span>
                                导出汇总
                            </button>
                            <button class="btn btn-secondary" id="exportDetailsBtn">
                                <span>📋</span>
                                导出明细
                            </button>
                        </div>
                    </div>
                    
                    <!-- 图表区域 -->
                    <div class="chart-container" id="chartContainer">
                        <canvas id="salesChart"></canvas>
                    </div>
                    
                    <!-- 数据表格 -->
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>月份</th>
                                    <th style="text-align: right;">订单数</th>
                                    <th style="text-align: right;">支付金额</th>
                                    <th style="text-align: right;">充值抵扣</th>
                                    <th style="text-align: right;">总金额</th>
                                    <th style="text-align: right;">环比增长</th>
                                </tr>
                            </thead>
                            <tbody id="resultTable"></tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- 加载遮罩 -->
            <div class="loading-overlay" id="loadingOverlay">
                <div class="loading-content">
                    <div class="spinner"></div>
                    <div class="loading-step" id="loadingStep">步骤 1/3</div>
                    <div class="loading-text" id="loadingText">正在处理...</div>
                    <div class="progress-bar">
                        <div class="progress-bar-fill" id="progressBarFill"></div>
                    </div>
                    <div class="loading-detail" id="loadingDetail"></div>
                    <div class="loading-timer" id="loadingTimer">已用时: 0秒</div>
                    <button class="btn btn-secondary" id="cancelBtn">取消</button>
                </div>
            </div>
            
            <style>
                .analysis-options { margin-bottom: 32px; }
                
                .options-card {
                    background: var(--bg-card);
                    border: 1px solid var(--border-color);
                    border-radius: 16px;
                    padding: 32px;
                }
                
                .options-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 24px;
                }
                
                .options-header h3 { color: var(--text-primary); margin: 0; }
                
                .cache-status {
                    font-size: 0.85rem;
                    color: var(--accent-green);
                    background: rgba(16, 185, 129, 0.1);
                    padding: 6px 12px;
                    border-radius: 20px;
                }
                
                .filter-row {
                    margin-bottom: 24px;
                }
                
                .filter-item label {
                    display: block;
                    margin-bottom: 12px;
                    color: var(--text-secondary);
                    font-size: 0.95rem;
                    font-weight: 500;
                }
                
                .option-tabs {
                    display: flex;
                    gap: 8px;
                    flex-wrap: wrap;
                }
                
                .option-tab {
                    padding: 10px 20px;
                    border: 1px solid var(--border-color);
                    background: var(--bg-secondary);
                    color: var(--text-secondary);
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 0.9rem;
                    transition: var(--transition-fast);
                }
                
                .option-tab:hover {
                    border-color: var(--accent-blue);
                    color: var(--text-primary);
                }
                
                .option-tab.active {
                    background: var(--gradient-primary);
                    color: white;
                    border-color: transparent;
                }
                
                .option-tab.disabled {
                    opacity: 0.5;
                    cursor: not-allowed;
                }
                
                .select-input {
                    width: 100%;
                    max-width: 400px;
                    padding: 12px 16px;
                    border: 1px solid var(--border-color);
                    border-radius: 8px;
                    background: var(--bg-secondary);
                    color: var(--text-primary);
                    font-size: 1rem;
                    cursor: pointer;
                }
                
                .select-input:focus {
                    outline: none;
                    border-color: var(--accent-blue);
                }
                
                .result-subtitle {
                    color: var(--text-muted);
                    font-size: 0.95rem;
                    margin-top: 8px;
                }
                
                .chart-container {
                    background: var(--bg-secondary);
                    border-radius: 12px;
                    padding: 24px;
                    margin-bottom: 24px;
                    height: 350px;
                }
                
                .growth-positive { color: var(--accent-green); }
                .growth-negative { color: var(--accent-red); }
                
                .data-source-checkboxes {
                    display: flex;
                    flex-direction: column;
                    gap: 12px;
                    max-height: 300px;
                    overflow-y: auto;
                    padding: 12px;
                    background: var(--bg-secondary);
                    border-radius: 8px;
                    border: 1px solid var(--border-color);
                }
                
                .data-source-checkbox-item {
                    display: flex;
                    align-items: center;
                    cursor: pointer;
                    padding: 12px;
                    border-radius: 8px;
                    transition: background 0.2s;
                }
                
                .data-source-checkbox-item:hover {
                    background: rgba(59, 130, 246, 0.1);
                }
                
                .ds-checkbox {
                    width: 18px;
                    height: 18px;
                    margin-right: 12px;
                    cursor: pointer;
                    accent-color: var(--accent-blue);
                }
                
                .ds-checkbox-label {
                    display: flex;
                    flex-direction: column;
                    gap: 4px;
                    flex: 1;
                }
                
                .ds-checkbox-label strong {
                    color: var(--text-primary);
                    font-size: 0.95rem;
                }
                
                .ds-checkbox-meta {
                    color: var(--text-muted);
                    font-size: 0.85rem;
                }
                
                .target-input-wrapper {
                    position: relative;
                    width: 100%;
                    max-width: 400px;
                }
                
                .target-dropdown {
                    position: absolute;
                    top: 100%;
                    left: 0;
                    right: 0;
                    max-height: 300px;
                    overflow-y: auto;
                    background: var(--bg-card);
                    border: 1px solid var(--border-color);
                    border-radius: 8px;
                    box-shadow: var(--shadow-lg);
                    z-index: 1000;
                    margin-top: 4px;
                }
                
                .dropdown-item {
                    padding: 12px 16px;
                    cursor: pointer;
                    color: var(--text-primary);
                    transition: background 0.2s;
                    border-bottom: 1px solid var(--border-color);
                }
                
                .dropdown-item:last-child {
                    border-bottom: none;
                }
                
                .dropdown-item:hover {
                    background: rgba(59, 130, 246, 0.1);
                }
                
                .dropdown-item:active {
                    background: rgba(59, 130, 246, 0.2);
                }
            </style>
        `;
        
        this.bindEvents(container);
        this.setupProgressListener();
        await this.checkDataSource();
    }
    
    async checkDataSource() {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        const uploadSection = document.getElementById('uploadSection');
        const dataSourceNotice = document.getElementById('dataSourceNotice');
        
        if (!uploadSection || !dataSourceNotice) {
            console.error('DOM元素未找到');
            return;
        }
        
        try {
            const listInfo = await invoke('get_data_source_list_info');
            
            if (listInfo && listInfo.data_sources && listInfo.data_sources.length > 0) {
                // 有数据源，显示分析选项
                uploadSection.style.display = 'block';
                dataSourceNotice.style.display = 'none';
                
                // 填充数据源checkbox列表
                const dataSourceCheckboxes = document.getElementById('dataSourceCheckboxes');
                dataSourceCheckboxes.innerHTML = listInfo.data_sources.map(ds => {
                    const checked = listInfo.current_id === ds.id ? 'checked' : '';
                    return `
                        <label class="data-source-checkbox-item">
                            <input type="checkbox" value="${ds.id}" ${checked} class="ds-checkbox">
                            <span class="ds-checkbox-label">
                                <strong>${this.escapeHtml(ds.file_name)}</strong>
                                <span class="ds-checkbox-meta">${ds.total_rows.toLocaleString()} 行</span>
                            </span>
                        </label>
                    `;
                }).join('');
                
                // 监听checkbox变化
                dataSourceCheckboxes.querySelectorAll('.ds-checkbox').forEach(checkbox => {
                    checkbox.addEventListener('change', () => {
                        this.updateDataSourceSelection();
                        this.updateSelectAllButton();
                    });
                });
                
                // 绑定全选按钮
                const selectAllBtn = document.getElementById('selectAllBtn');
                if (selectAllBtn) {
                    selectAllBtn.addEventListener('click', () => this.toggleSelectAll());
                }
                
                // 更新全选按钮状态
                this.updateSelectAllButton();
                
                // 如果有当前数据源，自动加载选项
                if (listInfo.current_id) {
                    try {
                        await invoke('auto_load_data_source');
                        await this.loadOptions();
                    } catch (error) {
                        console.error('自动加载数据源失败:', error);
                    }
                }
            } else {
                // 没有数据源，显示提示
                uploadSection.style.display = 'none';
                dataSourceNotice.style.display = 'block';
                document.getElementById('analysisOptions').style.display = 'none';
            }
        } catch (error) {
            console.error('检查数据源失败:', error);
            uploadSection.style.display = 'none';
            dataSourceNotice.style.display = 'block';
            document.getElementById('analysisOptions').style.display = 'none';
        }
    }
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    updateDataSourceSelection() {
        const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
        const selectedCount = checkboxes.length;
        
        if (selectedCount > 0) {
            this.loadOptions();
        } else {
            document.getElementById('analysisOptions').style.display = 'none';
            this.isDataLoaded = false;
        }
    }
    
    toggleSelectAll() {
        const checkboxes = document.querySelectorAll('.ds-checkbox');
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        
        checkboxes.forEach(checkbox => {
            checkbox.checked = !allChecked;
        });
        
        this.updateDataSourceSelection();
        this.updateSelectAllButton();
    }
    
    updateSelectAllButton() {
        const selectAllBtn = document.getElementById('selectAllBtn');
        if (!selectAllBtn) return;
        
        const checkboxes = document.querySelectorAll('.ds-checkbox');
        const allChecked = checkboxes.length > 0 && Array.from(checkboxes).every(cb => cb.checked);
        
        selectAllBtn.textContent = allChecked ? '取消全选' : '全选';
    }
    
    async loadOptions() {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            // 获取选中的数据源ID列表
            const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
            const selectedIds = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedIds.length === 0) {
                document.getElementById('analysisOptions').style.display = 'none';
                this.isDataLoaded = false;
                return;
            }
            
            let result;
            if (selectedIds.length === 1) {
                // 单个数据源，使用原有逻辑
                await invoke('switch_data_source', { dataSourceId: selectedIds[0] });
                result = await invoke('get_monthly_options');
            } else {
                // 多个数据源，使用合并选项
                result = await invoke('get_monthly_options_multi', { dataSourceIds: selectedIds });
            }
            
            if (result) {
                this.fileOptions = result;
                this.isDataLoaded = true;
                
                // 更新维度标签状态
                this.updateTabStates(result);
                
                // 初始化目标选择
                this.updateTargetSelect();
                
                document.getElementById('analysisOptions').style.display = 'block';
                const selectedCount = selectedIds.length;
                document.getElementById('cacheStatus').textContent = 
                    selectedCount > 1 ? `✅ ${selectedCount} 个数据源已合并` : '✅ 数据已缓存';
            }
        } catch (error) {
            console.error('加载选项失败:', error);
            this.showError('加载选项失败: ' + error);
        }
    }
    
    async setupProgressListener() {
        if (!window.__TAURI__) return;
        
        const { listen } = window.__TAURI__.event;
        
        if (this.unlistenProgress) {
            this.unlistenProgress();
        }
        
        this.unlistenProgress = await listen('monthly-progress', (event) => {
            this.updateProgress(event.payload);
        });
    }
    
    bindEvents(container) {
        const goToHomeBtn = container.querySelector('#goToHomeBtn');
        const cancelBtn = container.querySelector('#cancelBtn');
        const exportBtn = container.querySelector('#exportBtn');
        const exportDetailsBtn = container.querySelector('#exportDetailsBtn');
        const analyzeBtn = container.querySelector('#analyzeBtn');
        const optionTabs = container.querySelectorAll('.option-tab');
        
        if (goToHomeBtn) {
            goToHomeBtn.addEventListener('click', () => {
                window.location.hash = 'home';
            });
        }
        
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => this.cancelProcessing());
        }
        
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportResult());
        }
        
        if (exportDetailsBtn) {
            exportDetailsBtn.addEventListener('click', () => this.exportOrderDetails());
        }
        
        if (analyzeBtn) {
            analyzeBtn.addEventListener('click', () => this.runAnalysis());
        }
        
        optionTabs.forEach(tab => {
            tab.addEventListener('click', () => {
                if (tab.classList.contains('disabled')) return;
                optionTabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                this.currentDimension = tab.dataset.type;
                this.updateTargetSelect();
            });
        });
    }
    
    updateAnalyzeButton() {
        const targetInput = document.getElementById('targetInput');
        const analyzeBtn = document.getElementById('analyzeBtn');
        analyzeBtn.disabled = !targetInput?.value || targetInput.value.trim() === '';
    }
    
    // 格式化月份：将 "2024-01" 转换为 "2024年1月"
    formatMonth(monthStr) {
        if (!monthStr) return '未知月份';
        // 如果已经是中文格式，直接返回
        if (monthStr.includes('月')) return monthStr;
        
        // 如果是"未知月份"，返回原样
        if (monthStr === '未知月份') return monthStr;
        
        // 解析 "2024-01" 格式
        const match = monthStr.match(/^(\d{4})-(\d{1,2})$/);
        if (match) {
            const year = match[1];
            const month = parseInt(match[2], 10);
            return `${year}年${month}月`;
        }
        
        return monthStr;
    }
    
    updateTargetSelect() {
        const targetInput = document.getElementById('targetInput');
        const targetLabel = document.getElementById('targetLabel');
        const targetDropdown = document.getElementById('targetDropdown');
        
        if (!this.fileOptions) return;
        
        const labelMap = {
            'customer': '🎯 选择客户：',
            'province': '🎯 选择省份：',
            'city': '🎯 选择城市：',
            'district': '🎯 选择区县：',
            'region': '🎯 选择地区：',
        };
        
        targetLabel.textContent = labelMap[this.currentDimension] || labelMap['customer'];
        
        let options = [];
        switch (this.currentDimension) {
            case 'customer':
                options = this.fileOptions.available_customers.map(c => ({
                    value: c.code,
                    text: `${c.code} - ${c.name || '未知'}`,
                    searchText: `${c.code} ${c.name || ''}`.toLowerCase()
                }));
                break;
            case 'province':
                options = this.fileOptions.available_provinces.map(p => ({ 
                    value: p, 
                    text: p,
                    searchText: p.toLowerCase()
                }));
                break;
            case 'city':
                options = this.fileOptions.available_cities.map(c => ({ 
                    value: c, 
                    text: c,
                    searchText: c.toLowerCase()
                }));
                break;
            case 'district':
                options = this.fileOptions.available_districts.map(d => ({ 
                    value: d, 
                    text: d,
                    searchText: d.toLowerCase()
                }));
                break;
            case 'region':
                options = this.fileOptions.available_regions.map(r => ({ 
                    value: r, 
                    text: r,
                    searchText: r.toLowerCase()
                }));
                break;
        }
        
        // 保存选项供搜索使用
        this.currentOptions = options;
        
        // 清空输入框
        if (targetInput) {
            targetInput.value = '';
            targetInput.placeholder = '输入关键词搜索或选择...';
        }
        
        // 绑定输入事件
        if (targetInput && !targetInput.hasAttribute('data-bound')) {
            targetInput.setAttribute('data-bound', 'true');
            targetInput.addEventListener('input', (e) => this.handleTargetInput(e));
            targetInput.addEventListener('focus', () => this.showDropdown());
            targetInput.addEventListener('blur', () => {
                // 延迟隐藏，以便点击选项时能触发
                setTimeout(() => this.hideDropdown(), 200);
            });
        }
        
        this.updateAnalyzeButton();
    }
    
    handleTargetInput(e) {
        const query = e.target.value.trim().toLowerCase();
        const dropdown = document.getElementById('targetDropdown');
        
        if (!query) {
            this.showDropdown();
            return;
        }
        
        // 过滤选项
        const filtered = this.currentOptions.filter(opt => 
            opt.searchText.includes(query)
        );
        
        this.renderDropdown(filtered);
    }
    
    showDropdown() {
        const dropdown = document.getElementById('targetDropdown');
        if (!this.currentOptions || this.currentOptions.length === 0) return;
        
        // 如果有输入，显示过滤后的；否则显示全部
        const query = document.getElementById('targetInput')?.value.trim().toLowerCase() || '';
        const filtered = query 
            ? this.currentOptions.filter(opt => opt.searchText.includes(query))
            : this.currentOptions;
        
        this.renderDropdown(filtered);
    }
    
    hideDropdown() {
        const dropdown = document.getElementById('targetDropdown');
        if (dropdown) {
            dropdown.style.display = 'none';
        }
    }
    
    renderDropdown(options) {
        const dropdown = document.getElementById('targetDropdown');
        if (!dropdown) return;
        
        if (options.length === 0) {
            dropdown.innerHTML = '<div class="dropdown-item">无匹配结果</div>';
            dropdown.style.display = 'block';
            return;
        }
        
        dropdown.innerHTML = options.map(opt => 
            `<div class="dropdown-item" data-value="${this.escapeHtml(opt.value)}">${this.escapeHtml(opt.text)}</div>`
        ).join('');
        
        // 绑定点击事件
        dropdown.querySelectorAll('.dropdown-item').forEach(item => {
            item.addEventListener('click', (e) => {
                const value = e.target.dataset.value;
                const text = e.target.textContent;
                document.getElementById('targetInput').value = text;
                this.hideDropdown();
                this.updateAnalyzeButton();
            });
        });
        
        dropdown.style.display = 'block';
    }
    
    
    updateTabStates(result) {
        const tabs = document.querySelectorAll('.option-tab');
        const dataMap = {
            'customer': result.available_customers.length > 0,
            'province': result.available_provinces.length > 0,
            'city': result.available_cities.length > 0,
            'district': result.available_districts.length > 0,
            'region': result.available_regions.length > 0,
        };
        
        tabs.forEach(tab => {
            const type = tab.dataset.type;
            const hasData = dataMap[type] !== false;
            tab.classList.toggle('disabled', !hasData);
        });
    }
    
    async runAnalysis() {
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        if (!this.isDataLoaded) {
            this.showError('请先在首页导入数据源');
            await this.checkDataSource();
            return;
        }
        
        const targetInput = document.getElementById('targetInput');
        const targetValue = targetInput?.value || '';
        
        if (!targetValue || targetValue.trim() === '') {
            this.showError('请输入或选择分析目标');
            return;
        }
        
        // 从输入值中提取实际的值（如果是"编码 - 名称"格式，提取编码）
        let target = targetValue;
        if (this.currentDimension === 'customer') {
            const match = targetValue.match(/^([^\s-]+)/);
            if (match) {
                target = match[1];
            }
        }
        
        const activeTab = document.querySelector('.option-tab.active');
        const analysisType = activeTab?.dataset.type || 'customer';
        
        const analyzeBtn = document.getElementById('analyzeBtn');
        analyzeBtn.disabled = true;
        analyzeBtn.innerHTML = '<span>⏳</span> 分析中...';
        
        try {
            // 获取选中的数据源ID列表
            const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
            const selectedIds = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedIds.length === 0) {
                this.showError('请至少选择一个数据源');
                return;
            }
            
            let result;
            if (selectedIds.length === 1) {
                // 单个数据源，使用原有逻辑
                await invoke('switch_data_source', { dataSourceId: selectedIds[0] });
                result = await invoke('analyze_monthly_cached', {
                    analysisType,
                    target
                });
            } else {
                // 多个数据源，使用合并分析
                result = await invoke('analyze_monthly_multi', {
                    dataSourceIds: selectedIds,
                    analysisType,
                    target
                });
            }
            
            this.analysisResult = result;
            this.displayResult(result, analysisType);
            
        } catch (error) {
            this.showError('分析失败: ' + error);
        } finally {
            analyzeBtn.disabled = false;
            analyzeBtn.innerHTML = '<span>🔍</span> 开始分析';
            this.updateAnalyzeButton();
        }
    }
    
    displayResult(result, analysisType) {
        const typeTextMap = {
            'customer': '客户',
            'province': '省份',
            'city': '城市',
            'district': '区县',
            'region': '地区'
        };
        const typeText = typeTextMap[analysisType] || '维度';
        
        const targetName = result.target_name || result.target;
        document.getElementById('resultTitle').textContent = `${targetName} 月度销售趋势`;
        document.getElementById('resultSubtitle').textContent = 
            `${typeText}分析 · 共 ${result.monthly_data.length} 个月`;
        
        document.getElementById('totalAmount').textContent = 
            '¥' + result.total_amount.toLocaleString('zh-CN', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        document.getElementById('totalOrders').textContent = 
            result.total_orders.toLocaleString();
        document.getElementById('monthCount').textContent = 
            result.monthly_data.length;
        document.getElementById('processTime').textContent = 
            result.process_time_ms + 'ms';
        
        this.renderTable(result.monthly_data);
        this.renderChart(result.monthly_data);
        
        document.getElementById('resultSection').classList.add('visible');
        document.getElementById('resultSection').scrollIntoView({ behavior: 'smooth' });
        
        this.showToast(`✅ 分析完成！耗时 ${result.process_time_ms}ms`);
    }
    
    renderTable(data) {
        const tbody = document.getElementById('resultTable');
        tbody.innerHTML = '';
        
        data.forEach((item) => {
            const growthClass = item.mom_growth_rate > 0 ? 'growth-positive' : 
                               item.mom_growth_rate < 0 ? 'growth-negative' : '';
            const growthText = item.mom_growth_rate !== 0 
                ? `${item.mom_growth_rate > 0 ? '+' : ''}${item.mom_growth_rate.toFixed(2)}%`
                : '-';
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${this.formatMonth(item.month)}</td>
                <td style="text-align: right;">${item.order_count.toLocaleString()}</td>
                <td style="text-align: right;">
                    ¥${item.pay_amount.toLocaleString('zh-CN', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    })}
                </td>
                <td style="text-align: right;">
                    ¥${item.recharge_deduction.toLocaleString('zh-CN', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    })}
                </td>
                <td style="text-align: right;" class="amount">
                    ¥${item.total_amount.toLocaleString('zh-CN', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    })}
                </td>
                <td style="text-align: right;" class="${growthClass}">${growthText}</td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    renderChart(data) {
        const ctx = document.getElementById('salesChart').getContext('2d');
        const labels = data.map(d => this.formatMonth(d.month));
        const amounts = data.map(d => d.total_amount);
        
        this.drawLineChart(ctx, labels, amounts);
    }
    
    drawLineChart(ctx, labels, amounts) {
        const canvas = ctx.canvas;
        const width = canvas.offsetWidth;
        const height = canvas.offsetHeight;
        
        canvas.width = width * 2;
        canvas.height = height * 2;
        ctx.scale(2, 2);
        
        const padding = { top: 30, right: 30, bottom: 60, left: 80 };
        const chartWidth = width - padding.left - padding.right;
        const chartHeight = height - padding.top - padding.bottom;
        
        ctx.fillStyle = '#111827';
        ctx.fillRect(0, 0, width, height);
        
        if (amounts.length === 0) {
            ctx.fillStyle = '#94a3b8';
            ctx.font = '14px sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText('暂无数据', width / 2, height / 2);
            return;
        }
        
        const maxAmount = Math.max(...amounts) * 1.1 || 1;
        const minAmount = 0;
        
        // 网格线
        ctx.strokeStyle = '#2d3a4f';
        ctx.lineWidth = 1;
        
        for (let i = 0; i <= 5; i++) {
            const y = padding.top + (chartHeight / 5) * i;
            ctx.beginPath();
            ctx.moveTo(padding.left, y);
            ctx.lineTo(width - padding.right, y);
            ctx.stroke();
            
            const value = maxAmount - ((maxAmount - minAmount) / 5) * i;
            ctx.fillStyle = '#94a3b8';
            ctx.font = '11px sans-serif';
            ctx.textAlign = 'right';
            ctx.fillText('¥' + (value / 10000).toFixed(1) + 'w', padding.left - 10, y + 4);
        }
        
        // 折线
        const stepX = chartWidth / (labels.length - 1 || 1);
        
        // 渐变填充
        const gradient = ctx.createLinearGradient(0, padding.top, 0, height - padding.bottom);
        gradient.addColorStop(0, 'rgba(59, 130, 246, 0.3)');
        gradient.addColorStop(1, 'rgba(59, 130, 246, 0)');
        
        ctx.beginPath();
        ctx.moveTo(padding.left, height - padding.bottom);
        
        amounts.forEach((amount, i) => {
            const x = padding.left + i * stepX;
            const y = padding.top + chartHeight - ((amount - minAmount) / (maxAmount - minAmount)) * chartHeight;
            ctx.lineTo(x, y);
        });
        
        ctx.lineTo(padding.left + (labels.length - 1) * stepX, height - padding.bottom);
        ctx.closePath();
        ctx.fillStyle = gradient;
        ctx.fill();
        
        // 折线
        ctx.beginPath();
        ctx.strokeStyle = '#3b82f6';
        ctx.lineWidth = 3;
        
        amounts.forEach((amount, i) => {
            const x = padding.left + i * stepX;
            const y = padding.top + chartHeight - ((amount - minAmount) / (maxAmount - minAmount)) * chartHeight;
            if (i === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
        });
        ctx.stroke();
        
        // 数据点
        amounts.forEach((amount, i) => {
            const x = padding.left + i * stepX;
            const y = padding.top + chartHeight - ((amount - minAmount) / (maxAmount - minAmount)) * chartHeight;
            
            ctx.beginPath();
            ctx.arc(x, y, 5, 0, Math.PI * 2);
            ctx.fillStyle = '#3b82f6';
            ctx.fill();
            ctx.strokeStyle = '#fff';
            ctx.lineWidth = 2;
            ctx.stroke();
        });
        
        // X轴标签
        ctx.fillStyle = '#94a3b8';
        ctx.font = '10px sans-serif';
        ctx.textAlign = 'center';
        
        labels.forEach((label, i) => {
            const x = padding.left + i * stepX;
            ctx.fillText(label, x, height - padding.bottom + 20);
        });
        
        // 标题
        ctx.fillStyle = '#f0f4f8';
        ctx.font = '12px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('月度销售额趋势', width / 2, 20);
    }
    
    showLoading(step, text, percent, detail) {
        this.startTime = Date.now();
        document.getElementById('loadingOverlay').classList.add('visible');
        this.updateLoadingUI(step, text, percent, detail);
        
        if (this.timerInterval) clearInterval(this.timerInterval);
        this.timerInterval = setInterval(() => this.updateTimer(), 1000);
        this.updateTimer();
    }
    
    updateProgress(progress) {
        this.updateLoadingUI(`步骤 ${progress.step}`, progress.message, progress.percent, progress.detail);
    }
    
    updateLoadingUI(step, text, percent, detail) {
        document.getElementById('loadingStep').textContent = step;
        document.getElementById('loadingText').textContent = text;
        document.getElementById('progressBarFill').style.width = percent + '%';
        document.getElementById('loadingDetail').textContent = detail;
    }
    
    updateTimer() {
        if (this.startTime) {
            const elapsed = Math.floor((Date.now() - this.startTime) / 1000);
            const m = Math.floor(elapsed / 60);
            const s = elapsed % 60;
            document.getElementById('loadingTimer').textContent = `已用时: ${m > 0 ? m + '分' : ''}${s}秒`;
        }
    }
    
    hideLoading() {
        document.getElementById('loadingOverlay').classList.remove('visible');
        this.startTime = null;
        if (this.timerInterval) { clearInterval(this.timerInterval); this.timerInterval = null; }
    }
    
    async cancelProcessing() {
        if (!window.__TAURI__) {
            this.hideLoading();
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        try { await invoke('cancel_analysis'); } catch (e) {}
        this.hideLoading();
    }
    
    async exportResult() {
        if (!this.analysisResult || this.analysisResult.monthly_data.length === 0) {
            this.showError('没有数据可导出');
            return;
        }
        
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { save } = window.__TAURI__.dialog;
        
        const result = this.analysisResult;
        const headers = ['月份', '订单数', '支付金额', '充值抵扣', '总金额', '环比增长率'];
        const rows = result.monthly_data.map(item => [
            this.formatMonth(item.month),
            item.order_count,
            item.pay_amount.toFixed(2),
            item.recharge_deduction.toFixed(2),
            item.total_amount.toFixed(2),
            item.mom_growth_rate.toFixed(2) + '%'
        ]);
        
        // 添加BOM以支持中文
        let csvContent = '\uFEFF' + headers.join(',') + '\n';
        rows.forEach(row => {
            csvContent += row.map(cell => {
                const str = String(cell);
                if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                    return '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',') + '\n';
        });
        
        try {
            const filePath = await save({
                defaultPath: `月度销售分析汇总_${result.target_name || result.target}_${new Date().toISOString().slice(0,10)}.csv`,
                filters: [{
                    name: 'CSV文件',
                    extensions: ['csv']
                }]
            });
            
            if (filePath) {
                await invoke('save_export_file', {
                    filePath: filePath,
                    content: csvContent
                });
                
                this.showToast('✅ 导出成功！');
            }
        } catch (error) {
            console.error('导出失败:', error);
            if (error !== '用户取消操作') {
                this.showError('导出失败: ' + error);
            }
        }
    }
    
    async exportOrderDetails() {
        if (!this.analysisResult) {
            this.showError('请先进行分析');
            return;
        }
        
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { save } = window.__TAURI__.dialog;
        
        try {
            // 获取订单明细
            const activeTab = document.querySelector('.option-tab.active');
            const analysisType = activeTab?.dataset.type || 'customer';
            const targetInput = document.getElementById('targetInput');
            const targetValue = targetInput?.value || '';
            
            if (!targetValue || targetValue.trim() === '') {
                this.showError('请先输入或选择分析目标');
                return;
            }
            
            // 从输入值中提取实际的值
            let target = targetValue;
            if (analysisType === 'customer') {
                const match = targetValue.match(/^([^\s-]+)/);
                if (match) {
                    target = match[1];
                }
            }
            
            const orderDetails = await invoke('get_order_details', {
                analysisType,
                target
            });
            
            if (!orderDetails || orderDetails.length === 0) {
                this.showError('没有订单明细可导出');
                return;
            }
            
            // 生成CSV数据
            const headers = ['客户编码', '客户名称', '支付金额', '充值抵扣', '总金额', '省份', '城市', '区县', '地区', '月份'];
            const rows = orderDetails.map(row => [
                row.customer_code || '',
                row.customer_name || '',
                row.pay_amount.toFixed(2),
                row.recharge_deduction.toFixed(2),
                row.total_amount.toFixed(2),
                row.province || '',
                row.city || '',
                row.district || '',
                row.region || '',
                this.formatMonth(row.month) || ''
            ]);
            
            // 添加BOM以支持中文
            let csvContent = '\uFEFF' + headers.join(',') + '\n';
            rows.forEach(row => {
                csvContent += row.map(cell => {
                    const str = String(cell);
                    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                        return '"' + str.replace(/"/g, '""') + '"';
                    }
                    return str;
                }).join(',') + '\n';
            });
            
            const filePath = await save({
                defaultPath: `订单明细_${this.analysisResult.target_name || this.analysisResult.target}_${new Date().toISOString().slice(0,10)}.csv`,
                filters: [{
                    name: 'CSV文件',
                    extensions: ['csv']
                }]
            });
            
            if (filePath) {
                await invoke('save_export_file', {
                    filePath: filePath,
                    content: csvContent
                });
                
                this.showToast(`✅ 导出成功！共 ${orderDetails.length.toLocaleString()} 条订单明细`);
            }
        } catch (error) {
            console.error('导出订单明细失败:', error);
            if (error !== '用户取消操作') {
                this.showError('导出订单明细失败: ' + error);
            }
        }
    }
    
    showError(msg) { alert(msg); }
    
    showToast(message) {
        const existing = document.querySelector('.toast');
        if (existing) existing.remove();
        
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.style.cssText = `
            position: fixed; bottom: 40px; left: 50%; transform: translateX(-50%);
            background: var(--bg-card); border: 1px solid var(--accent-green);
            color: var(--text-primary); padding: 16px 32px; border-radius: 12px;
            box-shadow: var(--shadow-lg); z-index: 1000; animation: toastIn 0.3s ease;
        `;
        toast.textContent = message;
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'toastOut 0.3s ease forwards';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
}
