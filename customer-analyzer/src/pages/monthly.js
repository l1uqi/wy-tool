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
                            <span class="ds-info-label">当前数据源：</span>
                            <span class="ds-info-value" id="currentDataSource">-</span>
                        </div>
                    </div>
                    
                    <div class="info-box">
                        <h4>📌 所需列名说明</h4>
                        <ul>
                            <li><strong>客户编码</strong> - 用于识别唯一客户（必需）</li>
                            <li><strong>支付金额</strong> - 支付金额数值（必需）</li>
                            <li><strong>充值抵扣</strong> - 充值抵扣金额（必需）</li>
                            <li><strong>日期/下单日期</strong> - 用于按月汇总（推荐）</li>
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
                                <select id="targetSelect" class="select-input">
                                    <option value="">-- 请选择 --</option>
                                </select>
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
                        <button class="btn btn-primary" id="exportBtn">
                            <span>📥</span>
                            导出数据
                        </button>
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
            </style>
        `;
        
        this.bindEvents(container);
        this.setupProgressListener();
        await this.checkDataSource();
    }
    
    async checkDataSource() {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            const info = await invoke('get_data_source_info');
            if (info && info.file_path) {
                // 有数据源，显示分析选项
                document.getElementById('uploadSection').style.display = 'block';
                document.getElementById('dataSourceNotice').style.display = 'none';
                document.getElementById('currentDataSource').textContent = info.file_name || '未知文件';
                
                // 尝试自动加载数据源
                try {
                    await invoke('auto_load_data_source');
                    // 加载选项
                    await this.loadOptions();
                } catch (error) {
                    console.error('自动加载数据源失败:', error);
                }
            } else {
                // 没有数据源，显示提示
                document.getElementById('uploadSection').style.display = 'none';
                document.getElementById('dataSourceNotice').style.display = 'block';
                document.getElementById('analysisOptions').style.display = 'none';
            }
        } catch (error) {
            console.error('检查数据源失败:', error);
            document.getElementById('uploadSection').style.display = 'none';
            document.getElementById('dataSourceNotice').style.display = 'block';
            document.getElementById('analysisOptions').style.display = 'none';
        }
    }
    
    async loadOptions() {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            const result = await invoke('get_monthly_options');
            if (result) {
                this.fileOptions = result;
                this.isDataLoaded = true;
                
                // 更新维度标签状态
                this.updateTabStates(result);
                
                // 初始化目标选择
                this.updateTargetSelect();
                
                document.getElementById('analysisOptions').style.display = 'block';
                document.getElementById('cacheStatus').textContent = '✅ 数据已缓存';
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
        const analyzeBtn = container.querySelector('#analyzeBtn');
        const optionTabs = container.querySelectorAll('.option-tab');
        const targetSelect = container.querySelector('#targetSelect');
        
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
        
        if (analyzeBtn) {
            analyzeBtn.addEventListener('click', () => this.runAnalysis());
        }
        
        if (targetSelect) {
            targetSelect.addEventListener('change', () => this.updateAnalyzeButton());
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
        const targetSelect = document.getElementById('targetSelect');
        const analyzeBtn = document.getElementById('analyzeBtn');
        analyzeBtn.disabled = !targetSelect?.value;
    }
    
    updateTargetSelect() {
        const targetSelect = document.getElementById('targetSelect');
        const targetLabel = document.getElementById('targetLabel');
        
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
                    text: `${c.code} - ${c.name || '未知'}`
                }));
                break;
            case 'province':
                options = this.fileOptions.available_provinces.map(p => ({ value: p, text: p }));
                break;
            case 'city':
                options = this.fileOptions.available_cities.map(c => ({ value: c, text: c }));
                break;
            case 'district':
                options = this.fileOptions.available_districts.map(d => ({ value: d, text: d }));
                break;
            case 'region':
                options = this.fileOptions.available_regions.map(r => ({ value: r, text: r }));
                break;
        }
        
        targetSelect.innerHTML = '<option value="">-- 请选择 --</option>';
        options.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.text;
            targetSelect.appendChild(option);
        });
        
        this.updateAnalyzeButton();
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
        
        const targetSelect = document.getElementById('targetSelect');
        const target = targetSelect?.value || '';
        
        if (!target) {
            this.showError('请选择分析目标');
            return;
        }
        
        const activeTab = document.querySelector('.option-tab.active');
        const analysisType = activeTab?.dataset.type || 'customer';
        
        const analyzeBtn = document.getElementById('analyzeBtn');
        analyzeBtn.disabled = true;
        analyzeBtn.innerHTML = '<span>⏳</span> 分析中...';
        
        try {
            const result = await invoke('analyze_monthly_cached', {
                analysisType,
                target
            });
            
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
                <td>${item.month}</td>
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
        const labels = data.map(d => d.month);
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
    
    exportResult() {
        if (!this.analysisResult || this.analysisResult.monthly_data.length === 0) {
            this.showError('没有数据可导出');
            return;
        }
        
        const result = this.analysisResult;
        const headers = ['月份', '订单数', '支付金额', '充值抵扣', '总金额', '环比增长率'];
        const rows = result.monthly_data.map(item => [
            item.month,
            item.order_count,
            item.pay_amount.toFixed(2),
            item.recharge_deduction.toFixed(2),
            item.total_amount.toFixed(2),
            item.mom_growth_rate.toFixed(2) + '%'
        ]);
        
        let csv = '\uFEFF' + headers.join(',') + '\n';
        rows.forEach(row => { csv += row.join(',') + '\n'; });
        
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `月度销售分析_${result.target}.csv`;
        link.click();
        
        this.showToast('✅ 导出成功！');
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
