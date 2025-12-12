// 前20大客户分析页面
export class Top20Page {
    constructor(app) {
        this.app = app;
        this.processedData = [];
        this.totalAmountAll = 0;
        this.startTime = null;
        this.timerInterval = null;
        this.unlistenProgress = null;
        this.selectedDataSourceId = null;
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">📊</span>
                        前20大客户分析
                    </h1>
                    <p class="page-desc">
                        使用已导入的数据源进行分析，快速生成前20大客户排行榜
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
                        <button class="btn btn-primary" id="analyzeBtn" disabled>
                            <span>🔍</span>
                            开始分析
                        </button>
                    </div>
                    
                    <div class="info-box">
                        <h4>📌 所需列名说明</h4>
                        <ul>
                            <li><strong>客户编码</strong> - 用于识别唯一客户（必需）</li>
                            <li><strong>客户名称</strong> - 客户显示名称（可选）</li>
                            <li><strong>支付金额</strong> - 支付金额数值（必需）</li>
                            <li><strong>充值抵扣</strong> - 充值抵扣金额（必需）</li>
                        </ul>
                        <p>💡 金额计算公式：总金额 = 支付金额 + 充值抵扣</p>
                        <p style="margin-top: 12px; color: var(--accent-green);" id="dataSourceStatus">
                            ✅ 请选择数据源，点击"开始分析"按钮即可生成前20大客户排行榜
                        </p>
                    </div>
                </div>
                
                <div class="result-section" id="resultSection">
                    <div class="result-header">
                        <h2 class="result-title">
                            <span>🏆</span>
                            前20大客户排行
                        </h2>
                        <div class="stats-row">
                            <div class="stat-card">
                                <div class="stat-label">客户总数</div>
                                <div class="stat-value" id="totalCustomers">0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">总金额</div>
                                <div class="stat-value" id="totalAmount">¥0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">Top20占比</div>
                                <div class="stat-value" id="top20Percentage">0%</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">处理耗时</div>
                                <div class="stat-value" id="processTime">0ms</div>
                            </div>
                        </div>
                        <button class="btn btn-primary" id="exportBtn">
                            <span>📥</span>
                            导出结果
                        </button>
                    </div>
                    
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 70px; text-align: center;">排名</th>
                                    <th>客户编码</th>
                                    <th>客户名称</th>
                                    <th style="text-align: right;">订单数</th>
                                    <th style="text-align: right;">支付金额</th>
                                    <th style="text-align: right;">充值抵扣</th>
                                    <th style="text-align: right;">总金额</th>
                                    <th style="text-align: right;">占比</th>
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
                    <div class="loading-step" id="loadingStep">步骤 1/4</div>
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
            </style>
        `;
        
        this.bindEvents(container);
        this.setupProgressListener();
        // 确保DOM已经渲染后再检查数据源
        await this.checkDataSource();
    }
    
    async checkDataSource() {
        if (!window.__TAURI__) {
            console.warn('Tauri API 不可用');
            const uploadSection = document.getElementById('uploadSection');
            const dataSourceNotice = document.getElementById('dataSourceNotice');
            if (uploadSection) uploadSection.style.display = 'none';
            if (dataSourceNotice) dataSourceNotice.style.display = 'block';
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        await new Promise(resolve => setTimeout(resolve, 50));
        
        const uploadSection = document.getElementById('uploadSection');
        const dataSourceNotice = document.getElementById('dataSourceNotice');
        const dataSourceCheckboxes = document.getElementById('dataSourceCheckboxes');
        
        if (!uploadSection || !dataSourceNotice || !dataSourceCheckboxes) {
            console.error('DOM元素未找到');
            return;
        }
        
        try {
            const listInfo = await invoke('get_data_source_list_info');
            console.log('前20大客户分析 - 数据源列表:', listInfo);
            
            if (listInfo && listInfo.data_sources && listInfo.data_sources.length > 0) {
                // 有数据源，显示分析选项
                uploadSection.style.display = 'block';
                dataSourceNotice.style.display = 'none';
                
                // 填充数据源checkbox列表
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
                        this.updateAnalyzeButton();
                        this.updateSelectAllButton();
                    });
                });
                
                // 绑定全选按钮
                const selectAllBtn = document.getElementById('selectAllBtn');
                if (selectAllBtn) {
                    selectAllBtn.addEventListener('click', () => this.toggleSelectAll());
                }
                
                // 初始化按钮状态
                this.updateAnalyzeButton();
                this.updateSelectAllButton();
            } else {
                // 没有数据源，显示提示
                uploadSection.style.display = 'none';
                dataSourceNotice.style.display = 'block';
            }
        } catch (error) {
            console.error('前20大客户分析 - 检查数据源失败:', error);
            uploadSection.style.display = 'none';
            dataSourceNotice.style.display = 'block';
        }
    }
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    async setupProgressListener() {
        if (!window.__TAURI__) return;
        
        const { listen } = window.__TAURI__.event;
        
        // 监听Rust后端的进度事件
        if (this.unlistenProgress) {
            this.unlistenProgress();
        }
        
        this.unlistenProgress = await listen('excel-progress', (event) => {
            const progress = event.payload;
            this.updateProgress(progress);
        });
    }
    
    bindEvents(container) {
        const analyzeBtn = container.querySelector('#analyzeBtn');
        const goToHomeBtn = container.querySelector('#goToHomeBtn');
        const cancelBtn = container.querySelector('#cancelBtn');
        const exportBtn = container.querySelector('#exportBtn');
        
        if (analyzeBtn) {
            analyzeBtn.addEventListener('click', () => this.runAnalysis());
        }
        
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
    }
    
    updateAnalyzeButton() {
        const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const statusText = document.getElementById('dataSourceStatus');
        
        const selectedCount = checkboxes.length;
        
        if (analyzeBtn) {
            analyzeBtn.disabled = selectedCount === 0;
        }
        
        if (statusText) {
            if (selectedCount === 0) {
                statusText.textContent = '⚠️ 请至少选择一个数据源';
            } else if (selectedCount === 1) {
                const selected = checkboxes[0];
                const label = selected.closest('.data-source-checkbox-item').querySelector('strong').textContent;
                statusText.textContent = `✅ 已选择 1 个数据源：${label}，点击"开始分析"按钮即可生成前20大客户排行榜`;
            } else {
                statusText.textContent = `✅ 已选择 ${selectedCount} 个数据源，将合并分析，点击"开始分析"按钮即可生成前20大客户排行榜`;
            }
        }
    }
    
    toggleSelectAll() {
        const checkboxes = document.querySelectorAll('.ds-checkbox');
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        
        checkboxes.forEach(checkbox => {
            checkbox.checked = !allChecked;
        });
        
        this.updateAnalyzeButton();
        this.updateSelectAllButton();
    }
    
    updateSelectAllButton() {
        const selectAllBtn = document.getElementById('selectAllBtn');
        if (!selectAllBtn) return;
        
        const checkboxes = document.querySelectorAll('.ds-checkbox');
        const allChecked = checkboxes.length > 0 && Array.from(checkboxes).every(cb => cb.checked);
        
        selectAllBtn.textContent = allChecked ? '取消全选' : '全选';
    }
    
    async runAnalysis() {
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        // 获取选中的数据源ID列表
        const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
        const selectedIds = Array.from(checkboxes).map(cb => cb.value);
        
        if (selectedIds.length === 0) {
            this.showError('请至少选择一个数据源');
            return;
        }
        
        // 显示加载界面
        this.showLoading('步骤 1/1', '正在分析数据...', 0, '');
        
        try {
            let result;
            if (selectedIds.length === 1) {
                // 单个数据源，使用原有逻辑
                await invoke('switch_data_source', { dataSourceId: selectedIds[0] });
                result = await invoke('analyze_top20_cached');
            } else {
                // 多个数据源，使用合并分析
                result = await invoke('analyze_top20_multi', { dataSourceIds: selectedIds });
            }
            
            this.handleResult(result);
        } catch (error) {
            this.hideLoading();
            if (error !== '用户取消操作') {
                console.error('分析失败:', error);
                this.showError('分析失败: ' + error);
                
                // 如果是因为没有数据源，显示提示
                if (error.includes('数据源') || error.includes('导入')) {
                    await this.checkDataSource();
                }
            }
        }
    }
    
    handleResult(result) {
        this.processedData = result.top20;
        this.totalAmountAll = result.total_amount;
        
        // 更新统计信息
        document.getElementById('totalCustomers').textContent = 
            result.total_customers.toLocaleString();
        document.getElementById('totalAmount').textContent = 
            '¥' + result.total_amount.toLocaleString('zh-CN', { 
                minimumFractionDigits: 2, 
                maximumFractionDigits: 2 
            });
        document.getElementById('top20Percentage').textContent = 
            result.total_amount > 0 
                ? ((result.top20_amount / result.total_amount) * 100).toFixed(2) + '%'
                : '0%';
        document.getElementById('processTime').textContent = 
            result.process_time_ms + 'ms';
        
        // 渲染表格
        this.renderTable(result.top20);
        
        // 显示结果区域
        const resultSection = document.getElementById('resultSection');
        resultSection.classList.add('visible');
        resultSection.scrollIntoView({ behavior: 'smooth' });
        
        this.hideLoading();
        
        // 显示完成提示
        this.showToast(`✅ 处理完成！耗时 ${result.process_time_ms}ms，共 ${result.total_rows.toLocaleString()} 行数据`);
    }
    
    renderTable(data) {
        const tbody = document.getElementById('resultTable');
        tbody.innerHTML = '';
        
        data.forEach((customer, index) => {
            const rank = index + 1;
            const percentage = this.totalAmountAll > 0 
                ? ((customer.total_amount / this.totalAmountAll) * 100).toFixed(2)
                : 0;
            
            let rankClass = 'rank-other';
            if (rank === 1) rankClass = 'rank-1';
            else if (rank === 2) rankClass = 'rank-2';
            else if (rank === 3) rankClass = 'rank-3';
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align: center;">
                    <span class="rank-badge ${rankClass}">${rank}</span>
                </td>
                <td>${this.escapeHtml(customer.customer_code)}</td>
                <td class="customer-name" title="${this.escapeHtml(customer.customer_name)}">
                    ${this.escapeHtml(customer.customer_name) || '-'}
                </td>
                <td style="text-align: right;">${customer.order_count.toLocaleString()}</td>
                <td style="text-align: right;">
                    ¥${customer.pay_amount.toLocaleString('zh-CN', { 
                        minimumFractionDigits: 2, 
                        maximumFractionDigits: 2 
                    })}
                </td>
                <td style="text-align: right;">
                    ¥${customer.recharge_deduction.toLocaleString('zh-CN', { 
                        minimumFractionDigits: 2, 
                        maximumFractionDigits: 2 
                    })}
                </td>
                <td style="text-align: right;" class="amount">
                    ¥${customer.total_amount.toLocaleString('zh-CN', { 
                        minimumFractionDigits: 2, 
                        maximumFractionDigits: 2 
                    })}
                </td>
                <td style="text-align: right;" class="percentage">${percentage}%</td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    showLoading(step, text, percent, detail) {
        this.startTime = Date.now();
        
        const overlay = document.getElementById('loadingOverlay');
        overlay.classList.add('visible');
        
        this.updateLoadingUI(step, text, percent, detail);
        
        // 启动计时器
        if (this.timerInterval) {
            clearInterval(this.timerInterval);
        }
        this.timerInterval = setInterval(() => this.updateTimer(), 1000);
        this.updateTimer();
    }
    
    updateProgress(progress) {
        this.updateLoadingUI(
            `步骤 ${progress.step}`,
            progress.message,
            progress.percent,
            progress.detail
        );
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
            const minutes = Math.floor(elapsed / 60);
            const seconds = elapsed % 60;
            const timeStr = minutes > 0 ? `${minutes}分${seconds}秒` : `${seconds}秒`;
            document.getElementById('loadingTimer').textContent = `已用时: ${timeStr}`;
        }
    }
    
    hideLoading() {
        document.getElementById('loadingOverlay').classList.remove('visible');
        this.startTime = null;
        
        if (this.timerInterval) {
            clearInterval(this.timerInterval);
            this.timerInterval = null;
        }
    }
    
    async cancelProcessing() {
        if (!window.__TAURI__) {
            this.hideLoading();
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            await invoke('cancel_analysis');
        } catch (error) {
            console.error('取消失败:', error);
        }
        this.hideLoading();
    }
    
    async exportResult() {
        if (this.processedData.length === 0) {
            this.showError('没有数据可导出');
            return;
        }
        
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { save } = window.__TAURI__.dialog;
        
        // 生成CSV数据
        const headers = ['排名', '客户编码', '客户名称', '订单数', '支付金额', '充值抵扣', '总金额', '占比'];
        const rows = this.processedData.map((customer, index) => [
            index + 1,
            customer.customer_code,
            customer.customer_name,
            customer.order_count,
            customer.pay_amount.toFixed(2),
            customer.recharge_deduction.toFixed(2),
            customer.total_amount.toFixed(2),
            this.totalAmountAll > 0 
                ? ((customer.total_amount / this.totalAmountAll) * 100).toFixed(2) + '%'
                : '0%'
        ]);
        
        // 添加BOM以支持中文
        let csvContent = '\uFEFF' + headers.join(',') + '\n';
        rows.forEach(row => {
            csvContent += row.map(cell => {
                // 如果包含逗号或引号，需要用引号包裹
                const str = String(cell);
                if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                    return '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',') + '\n';
        });
        
        try {
            // 打开保存文件对话框
            const filePath = await save({
                defaultPath: `Top20客户分析结果_${new Date().toISOString().slice(0,10)}.csv`,
                filters: [{
                    name: 'CSV文件',
                    extensions: ['csv']
                }]
            });
            
            if (filePath) {
                // 调用后端命令保存文件
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
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    showError(message) {
        alert(message);
    }
    
    showToast(message) {
        const existingToast = document.querySelector('.toast');
        if (existingToast) {
            existingToast.remove();
        }
        
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.style.cssText = `
            position: fixed;
            bottom: 40px;
            left: 50%;
            transform: translateX(-50%);
            background: var(--bg-card);
            border: 1px solid var(--accent-green);
            color: var(--text-primary);
            padding: 16px 32px;
            border-radius: 12px;
            box-shadow: var(--shadow-lg), 0 0 20px rgba(16, 185, 129, 0.2);
            z-index: 1000;
            animation: toastIn 0.3s ease;
        `;
        toast.textContent = message;
        
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'toastOut 0.3s ease forwards';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
}

