// 客户采购额计算页面
export class PurchasePage {
    constructor(app) {
        this.app = app;
        this.selectedDataSourceIds = [];
        this.customerCodes = [];
        this.resultData = [];
        this.originalExcelData = null; // 存储原始Excel数据
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">💰</span>
                        客户采购额计算
                    </h1>
                    <p class="page-desc">
                        导入客户编码表，自动匹配数据源中的客户数据，按月份统计每个客户的采购金额
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
                    <!-- 数据源选择 -->
                    <div class="data-source-info-card">
                        <div class="ds-info-header">
                            <div class="ds-select-group">
                                <label class="ds-info-label">选择数据源（可多选合并分析）：</label>
                                <div class="data-source-checkboxes" id="dataSourceCheckboxes">
                                    <p style="color: var(--text-muted);">加载中...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- 客户编码文件上传 -->
                    <div class="upload-group">
                            <div>
                            <div style="margin-bottom: 8px; font-weight: bold;">导入客户编码表：</div>
                            <div id="btnCustomerFile" class="btn btn-primary">
                                <span>📥</span>
                                点击上传客户编码表（Excel）
                            </div>
                            <div id="customerFileInfo" class="file-info" style="margin-top: 8px; display: none;">
                                <span id="customerFileName"></span>
                                <span id="customerFileCount"></span>
                            </div>
                        </div>
                        <div class="info-hint">
                            <p>💡 Excel文件需包含"客户编码"列，系统将根据编码匹配数据源中的客户数据</p>
                        </div>
                    </div>
                    
                    <button class="btn btn-primary" id="analyzeBtn" disabled>
                        <span>🔍</span>
                        开始计算
                    </button>
                </div>
                
                <!-- 结果区域 -->
                <div class="result-section" id="resultSection" style="display: none;">
                    <div class="result-header">
                        <div>
                            <h2 class="result-title">
                                <span>📊</span>
                                客户采购额统计结果
                            </h2>
                            <p class="result-subtitle" id="resultSubtitle"></p>
                        </div>
                        <div class="stats-row">
                            <div class="stat-card">
                                <div class="stat-label">客户总数</div>
                                <div class="stat-value" id="totalCustomers">0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">总采购额</div>
                                <div class="stat-value" id="totalAmount">¥0</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-label">月份数</div>
                                <div class="stat-value" id="monthCount">0</div>
                            </div>
                        </div>
                        <button class="btn btn-primary" id="exportBtn">
                            <span>📥</span>
                            导出结果
                        </button>
                    </div>
                    
                    <!-- 数据表格 -->
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>客户编码</th>
                                    <th>客户名称</th>
                                    <th id="monthHeaders"></th>
                                    <th style="text-align: right;">合计</th>
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
                    <button class="btn btn-secondary" id="cancelBtn">取消</button>
                </div>
            </div>
            
            <style>
                .upload-group {
                    background: var(--bg-card);
                    padding: 24px;
                    border-radius: 12px;
                    border: 1px solid var(--border-color);
                    margin-bottom: 24px;
                }
                
                .file-info {
                    color: var(--text-secondary);
                    font-size: 0.9rem;
                    display: flex;
                    flex-direction: column;
                    gap: 4px;
                }
                
                .info-hint {
                    margin-top: 12px;
                    padding: 12px;
                    background: rgba(59, 130, 246, 0.1);
                    border-radius: 8px;
                    color: var(--text-secondary);
                    font-size: 0.9rem;
                }
                
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
                uploadSection.style.display = 'block';
                dataSourceNotice.style.display = 'none';
                
                // 填充数据源checkbox列表
                const dataSourceCheckboxes = document.getElementById('dataSourceCheckboxes');
                dataSourceCheckboxes.innerHTML = listInfo.data_sources.map(ds => {
                    return `
                        <label class="data-source-checkbox-item">
                            <input type="checkbox" value="${ds.id}" class="ds-checkbox">
                            <span class="ds-checkbox-label">
                                <strong>${this.escapeHtml(ds.file_name)}</strong>
                                <span class="ds-checkbox-meta">${ds.total_rows.toLocaleString()} 行</span>
                            </span>
                        </label>
                    `;
                }).join('');
                
                // 监听checkbox变化
                dataSourceCheckboxes.querySelectorAll('.ds-checkbox').forEach(checkbox => {
                    checkbox.addEventListener('change', () => this.updateAnalyzeButton());
                });
                
                this.updateAnalyzeButton();
            } else {
                uploadSection.style.display = 'none';
                dataSourceNotice.style.display = 'block';
            }
        } catch (error) {
            console.error('检查数据源失败:', error);
            uploadSection.style.display = 'none';
            dataSourceNotice.style.display = 'block';
        }
    }
    
    bindEvents(container) {
        const goToHomeBtn = container.querySelector('#goToHomeBtn');
        const btnCustomerFile = container.querySelector('#btnCustomerFile');
        const analyzeBtn = container.querySelector('#analyzeBtn');
        const exportBtn = container.querySelector('#exportBtn');
        const cancelBtn = container.querySelector('#cancelBtn');
        
        if (goToHomeBtn) {
            goToHomeBtn.addEventListener('click', () => {
                window.location.hash = 'home';
            });
        }
        
        if (btnCustomerFile) {
            btnCustomerFile.addEventListener('click', () => {
                this.handleCustomerFile();
            });
        }
        
        if (analyzeBtn) {
            analyzeBtn.addEventListener('click', () => this.runAnalysis());
        }
        
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportResult());
        }
        
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => this.cancelProcessing());
        }
    }
    
    async handleCustomerFile() {
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { open } = window.__TAURI__.dialog;
        
        try {
            // 使用Tauri的文件选择对话框
            const selected = await open({
                multiple: false,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx', 'xls']
                }]
            });
            
            if (!selected) {
                return;
            }
            
            this.showLoading('步骤 1/2', '正在读取客户编码文件...', 0, '');
            
            // 读取文件并提取客户编码和完整数据
            const result = await invoke('load_customer_codes', { filePath: selected });
            
            this.customerCodes = result.customer_codes || [];
            // 保存原始Excel数据
            this.originalExcelData = {
                headers: result.headers || [],
                rows: result.rows || [],
                customerCodeIndex: result.customer_code_index || 0
            };
            
            this.hideLoading();
            
            // 更新UI
            const fileName = selected.split(/[/\\]/).pop() || selected;
            document.getElementById('customerFileName').textContent = `文件：${fileName}`;
            document.getElementById('customerFileCount').textContent = `共 ${this.customerCodes.length} 个客户编码`;
            document.getElementById('customerFileInfo').style.display = 'block';
            
            this.updateAnalyzeButton();
            
            this.showToast(`✅ 成功导入 ${this.customerCodes.length} 个客户编码`);
        } catch (error) {
            this.hideLoading();
            if (error !== '用户取消操作') {
                this.showError('读取客户编码文件失败: ' + error);
            }
        }
    }
    
    updateAnalyzeButton() {
        const checkboxes = document.querySelectorAll('.ds-checkbox:checked');
        const selectedCount = checkboxes.length;
        const hasCustomerCodes = this.customerCodes.length > 0;
        const analyzeBtn = document.getElementById('analyzeBtn');
        
        if (analyzeBtn) {
            analyzeBtn.disabled = selectedCount === 0 || !hasCustomerCodes;
        }
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
        
        if (this.customerCodes.length === 0) {
            this.showError('请先导入客户编码表');
            return;
        }
        
        try {
            this.showLoading('步骤 1/2', '正在计算客户采购额...', 0, '');
            
            const result = await invoke('calculate_customer_purchase', {
                dataSourceIds: selectedIds,
                customerCodes: this.customerCodes
            });
            
            this.resultData = result;
            
            this.hideLoading();
            
            // 直接导出，不显示预览
            await this.exportResult();
        } catch (error) {
            this.hideLoading();
            this.showError('计算失败: ' + error);
        }
    }
    
    displayResult(result) {
        document.getElementById('totalCustomers').textContent = result.total_customers.toLocaleString();
        document.getElementById('totalAmount').textContent = 
            '¥' + result.total_amount.toLocaleString('zh-CN', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        
        // 获取所有月份并排序
        const months = new Set();
        result.customer_data.forEach(customer => {
            customer.monthly_data.forEach(m => months.add(m.month));
        });
        const sortedMonths = Array.from(months).sort();
        document.getElementById('monthCount').textContent = sortedMonths.length;
        
        // 生成月份表头
        const monthHeaders = document.getElementById('monthHeaders');
        monthHeaders.innerHTML = sortedMonths.map(month => 
            `<th style="text-align: right;">${this.formatMonth(month)}</th>`
        ).join('');
        
        // 渲染表格
        this.renderTable(result.customer_data, sortedMonths);
        
        document.getElementById('resultSection').style.display = 'block';
        document.getElementById('resultSection').scrollIntoView({ behavior: 'smooth' });
    }
    
    renderTable(customerData, months) {
        const tbody = document.getElementById('resultTable');
        tbody.innerHTML = '';
        
        customerData.forEach(customer => {
            const tr = document.createElement('tr');
            
            // 创建月份数据映射
            const monthlyMap = new Map();
            customer.monthly_data.forEach(m => {
                monthlyMap.set(m.month, m.total_amount);
            });
            
            // 计算合计
            const total = customer.monthly_data.reduce((sum, m) => sum + m.total_amount, 0);
            
            // 构建行HTML
            let rowHtml = `
                <td>${this.escapeHtml(customer.customer_code)}</td>
                <td>${this.escapeHtml(customer.customer_name || '-')}</td>
            `;
            
            // 添加每个月的金额
            months.forEach(month => {
                const amount = monthlyMap.get(month) || 0;
                rowHtml += `
                    <td style="text-align: right;">
                        ${amount > 0 ? '¥' + amount.toLocaleString('zh-CN', {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        }) : '-'}
                    </td>
                `;
            });
            
            // 添加合计
            rowHtml += `
                <td style="text-align: right;" class="amount">
                    ¥${total.toLocaleString('zh-CN', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    })}
                </td>
            `;
            
            tr.innerHTML = rowHtml;
            tbody.appendChild(tr);
        });
    }
    
    formatMonth(monthStr) {
        if (!monthStr) return monthStr;
        if (monthStr.includes('月')) return monthStr;
        
        const match = monthStr.match(/^(\d{4})-(\d{1,2})$/);
        if (match) {
            const month = parseInt(match[2], 10);
            return `${month}月`;
        }
        
        return monthStr;
    }
    
    // 按时间顺序排序月份
    sortMonths(months) {
        return months.sort((a, b) => {
            // 如果是"未知月份"，排在最后
            if (a === '未知月份') return 1;
            if (b === '未知月份') return -1;
            
            // 解析 "2024-01" 格式
            const matchA = a.match(/^(\d{4})-(\d{1,2})$/);
            const matchB = b.match(/^(\d{4})-(\d{1,2})$/);
            
            if (matchA && matchB) {
                const yearA = parseInt(matchA[1], 10);
                const monthA = parseInt(matchA[2], 10);
                const yearB = parseInt(matchB[1], 10);
                const monthB = parseInt(matchB[2], 10);
                
                if (yearA !== yearB) {
                    return yearA - yearB;
                }
                return monthA - monthB;
            }
            
            // 如果格式不匹配，使用字符串排序
            return a.localeCompare(b);
        });
    }
    
    async exportResult() {
        if (!this.resultData || !this.originalExcelData) {
            this.showError('没有数据可导出');
            return;
        }
        
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { save } = window.__TAURI__.dialog;
        
        try {
            // 获取所有月份并排序
            const months = new Set();
            this.resultData.customer_data.forEach(customer => {
                customer.monthly_data.forEach(m => months.add(m.month));
            });
            const sortedMonths = this.sortMonths(Array.from(months));
            
            // 创建客户数据映射（用于快速查找）
            const customerDataMap = new Map();
            this.resultData.customer_data.forEach(customer => {
                customerDataMap.set(customer.customer_code, customer);
            });
            
            // 生成表头：原始Excel的所有列 + 月份列 + 合计列
            const headers = [
                ...this.originalExcelData.headers,
                ...sortedMonths.map(m => this.formatMonth(m)),
                '合计'
            ];
            
            // 生成数据行：保留原始Excel的所有列，然后添加月份数据和合计
            const rows = this.originalExcelData.rows.map((originalRow) => {
                const code = originalRow[this.originalExcelData.customerCodeIndex] || '';
                const customer = customerDataMap.get(code);
                
                // 创建月份数据映射
                const monthlyMap = new Map();
                let total = 0;
                
                if (customer) {
                    customer.monthly_data.forEach(m => {
                        monthlyMap.set(m.month, m.total_amount);
                    });
                    total = customer.monthly_data.reduce((sum, m) => sum + m.total_amount, 0);
                }
                
                // 构建行：原始列 + 月份列 + 合计
                return [
                    ...originalRow,
                    ...sortedMonths.map(month => {
                        const amount = monthlyMap.get(month) || 0;
                        return amount > 0 ? amount.toFixed(2) : '';
                    }),
                    total > 0 ? total.toFixed(2) : '0.00'
                ];
            });
            
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
                defaultPath: `客户采购额统计_${new Date().toISOString().slice(0,10)}.csv`,
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
    
    showLoading(step, text, percent, detail) {
        const overlay = document.getElementById('loadingOverlay');
        overlay.classList.add('visible');
        document.getElementById('loadingStep').textContent = step;
        document.getElementById('loadingText').textContent = text;
        document.getElementById('progressBarFill').style.width = percent + '%';
        document.getElementById('loadingDetail').textContent = detail;
    }
    
    hideLoading() {
        const overlay = document.getElementById('loadingOverlay');
        overlay.classList.remove('visible');
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
            box-shadow: var(--shadow-lg);
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

