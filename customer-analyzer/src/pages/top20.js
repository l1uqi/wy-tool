// 前20大客户分析页面
const { invoke } = window.__TAURI__.core;
const { open } = window.__TAURI__.dialog;
const { listen } = window.__TAURI__.event;

export class Top20Page {
    constructor(app) {
        this.app = app;
        this.processedData = [];
        this.totalAmountAll = 0;
        this.startTime = null;
        this.timerInterval = null;
        this.unlistenProgress = null;
    }
    
    render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">📊</span>
                        前20大客户分析
                    </h1>
                    <p class="page-desc">
                        上传 Excel 文件，使用 Rust 高性能引擎快速分析客户数据
                    </p>
                </div>
                
                <div class="upload-section slide-up">
                    <div class="upload-area" id="uploadArea">
                        <div class="upload-icon">📁</div>
                        <div class="upload-text">点击选择 Excel 文件</div>
                        <div class="upload-hint">支持 .xlsx, .xls 格式（支持大文件，百万级数据）</div>
                    </div>
                    <div class="file-info" id="fileInfo"></div>
                    
                    <div class="info-box">
                        <h4>📌 所需列名说明</h4>
                        <ul>
                            <li><strong>客户编码</strong> - 用于识别唯一客户（必需）</li>
                            <li><strong>客户名称</strong> - 客户显示名称（可选）</li>
                            <li><strong>支付金额</strong> - 支付金额数值（必需）</li>
                            <li><strong>充值抵扣</strong> - 充值抵扣金额（必需）</li>
                        </ul>
                        <p>💡 金额计算公式：总金额 = 支付金额 + 充值抵扣</p>
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
        `;
        
        this.bindEvents(container);
        this.setupProgressListener();
    }
    
    async setupProgressListener() {
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
        const uploadArea = container.querySelector('#uploadArea');
        const cancelBtn = container.querySelector('#cancelBtn');
        const exportBtn = container.querySelector('#exportBtn');
        
        uploadArea.addEventListener('click', () => this.selectFile());
        cancelBtn.addEventListener('click', () => this.cancelProcessing());
        exportBtn.addEventListener('click', () => this.exportResult());
        
        // 拖拽支持（虽然Tauri桌面应用可能用不上，但保留兼容性）
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });
    }
    
    async selectFile() {
        try {
            const selected = await open({
                multiple: false,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx', 'xls']
                }]
            });
            
            if (selected) {
                await this.processFile(selected);
            }
        } catch (error) {
            console.error('选择文件失败:', error);
            this.showError('选择文件失败: ' + error);
        }
    }
    
    async processFile(filePath) {
        const fileName = filePath.split(/[/\\]/).pop();
        
        // 显示文件信息
        const fileInfo = document.getElementById('fileInfo');
        fileInfo.innerHTML = `📄 已选择文件: <strong>${fileName}</strong>`;
        fileInfo.classList.add('visible');
        
        // 显示加载界面
        this.showLoading('步骤 1/4', '准备处理文件...', 0, '');
        
        try {
            // 调用Rust后端处理Excel
            const result = await invoke('analyze_excel', { filePath });
            
            this.handleResult(result);
        } catch (error) {
            this.hideLoading();
            if (error !== '用户取消操作') {
                console.error('处理失败:', error);
                this.showError('处理失败: ' + error);
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
        try {
            await invoke('cancel_analysis');
        } catch (error) {
            console.error('取消失败:', error);
        }
        this.hideLoading();
    }
    
    exportResult() {
        if (this.processedData.length === 0) {
            this.showError('没有数据可导出');
            return;
        }
        
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
        
        // 创建下载
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Top20客户分析结果_${new Date().toISOString().slice(0,10)}.csv`;
        link.click();
        
        this.showToast('✅ 导出成功！');
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

