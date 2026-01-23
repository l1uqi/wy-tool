// 担保台账页面
import flatpickr from 'flatpickr';
import 'flatpickr/dist/flatpickr.min.css';
import 'flatpickr/dist/l10n/zh.js';

export class GuaranteePage {
    constructor(app) {
        this.app = app;
        this.data = this.loadData();
        this.currentEditId = null;
        this.currentView = 'list'; // 'list' 或 'form'
        this.sortColumn = null;
        this.sortDirection = 'asc';
        this.datePickerInstances = {};
        this.reserveData = this.loadReserveData(); // 预留数据
        this.filterGuarantor = '';
        this.filterGuaranteedCustomer = '';
        this.filterCustomer = '';
    }
    
    loadData() {
        try {
            const saved = localStorage.getItem('guarantee_records');
            return saved ? JSON.parse(saved) : [];
        } catch (error) {
            console.error('加载数据失败:', error);
            return [];
        }
    }
    
    loadReserveData() {
        try {
            const saved = localStorage.getItem('guarantee_reserve_data');
            return saved ? JSON.parse(saved) : [];
        } catch (error) {
            console.error('加载预留数据失败:', error);
            return [];
        }
    }
    
    saveReserveData() {
        try {
            localStorage.setItem('guarantee_reserve_data', JSON.stringify(this.reserveData));
        } catch (error) {
            console.error('保存预留数据失败:', error);
            this.showError('保存预留数据失败: ' + error);
        }
    }
    
    getTodayDate() {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
    
    saveData() {
        try {
            localStorage.setItem('guarantee_records', JSON.stringify(this.data));
        } catch (error) {
            console.error('保存数据失败:', error);
            this.showError('保存数据失败: ' + error);
        }
    }
    
    render(container) {
        if (this.currentView === 'form') {
            this.renderForm(container);
        } else {
            this.renderList(container);
        }
    }
    
    renderList(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up" style="display: none;">
                    <h1 class="page-title" style="font-size: 1.75rem;">
                        <span class="icon">📋</span>
                        担保台账
                    </h1>
                    <p class="page-desc">
                        记录和管理担保明细信息，支持手动录入和编辑
                    </p>
                </div>
                
                <div class="guarantee-container">
                    <!-- 数据表格区域 -->
                    <div class="guarantee-table-section slide-up">
                        <div class="table-card">
                            <div class="table-header">
                                <h3>担保明细列表</h3>
                                <div class="table-actions">
                                    <button class="btn btn-sm btn-primary" id="addNewBtn">
                                        ➕ 新增
                                    </button>
                                    <span class="record-count">共 ${this.data.length} 条记录</span>
                                    <button class="btn btn-sm btn-primary" id="importReserveBtn">
                                        导入预留数据
                                    </button>
                                    <button class="btn btn-sm btn-primary" id="importHistoryBtn">
                                        导入历史数据
                                    </button>
                                    <button class="btn btn-sm btn-primary" id="exportBtn" ${this.data.length === 0 ? 'disabled' : ''}>
                                        导出数据
                                    </button>
                                </div>
                            </div>
                            
                            <!-- 筛选功能 -->
                            <div class="filter-section">
                                <div class="filter-controls">
                                    <div class="filter-group">
                                        <label>担保方：</label>
                                        <input type="text" id="filterGuarantor" placeholder="输入担保方名称">
                                    </div>
                                    <div class="filter-group">
                                        <label>担保客户：</label>
                                        <input type="text" id="filterGuaranteedCustomer" placeholder="输入担保客户名称">
                                    </div>
                                    <div class="filter-group">
                                        <label>客户名称：</label>
                                        <input type="text" id="filterCustomer" placeholder="输入客户名称">
                                    </div>
                                    <div class="filter-actions">
                                        <button class="btn btn-sm btn-primary" id="applyFilterBtn">筛选</button>
                                        <button class="btn btn-sm btn-secondary" id="clearFilterBtn">清除</button>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="table-wrapper">
                                <table class="guarantee-table" id="guaranteeTable">
                                    <thead>
                                        <tr>
                                            <th class="sortable" data-column="registerTime" style="width: 140px; min-width: 140px;">登记时间 <span class="sort-icon">▼</span><div class="resizer"></div></th>
                                            <th class="sortable" data-column="guarantor" style="width: 100px; min-width: 100px;">担保方 <span class="sort-icon">▼</span><div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px;">担保类型<div class="resizer"></div></th>
                                            <th style="width: 100px; min-width: 100px;">对接人<div class="resizer"></div></th>
                                            <th style="width: 100px; min-width: 100px;">大区<div class="resizer"></div></th>
                                            <th style="width: 80px; min-width: 80px;">省区<div class="resizer"></div></th>
                                            <th style="width: 90px; min-width: 90px;">离职标识<div class="resizer"></div></th>
                                            <th style="width: 140px; min-width: 140px;">工资提成担保金额<div class="resizer"></div></th>
                                            <th style="width: 100px; min-width: 100px;">预留金额<div class="resizer"></div></th>
                                            <th style="width: 100px; min-width: 100px;">担保金额<div class="resizer"></div></th>
                                            <th style="width: 130px; min-width: 130px;">剩余可担保金额<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px;">担保客户<div class="resizer"></div></th>
                                            <th style="width: 110px; min-width: 110px; max-width: 110px;">订单号<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px; max-width: 120px;">预计回款时间<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px;">审批编号<div class="resizer"></div></th>
                                            <th style="width: 150px; min-width: 150px;">备注<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px;">工资提成回款<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px;">预留回款金额<div class="resizer"></div></th>
                                            <th style="width: 100px; min-width: 100px;">未回款金额<div class="resizer"></div></th>
                                            <th style="width: 120px; min-width: 120px; max-width: 120px;">回款时间<div class="resizer"></div></th>
                                            <th class="actions-col" style="width: 100px; min-width: 100px;">操作<div class="resizer"></div></th>
                                        </tr>
                                    </thead>
                                    <tbody id="guaranteeTableBody">
                                        ${this.renderTableBody()}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        this.bindListEvents(container);
    }
    
    renderForm(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title" style="font-size: 1.75rem;">
                        <span class="icon">📋</span>
                        担保台账
                    </h1>
                    <p class="page-desc">
                        ${this.currentEditId ? '编辑担保记录' : '新增担保记录'}
                    </p>
                </div>
                
                <div class="guarantee-container">
                    <!-- 录入表单区域 -->
                    <div class="guarantee-form-section slide-up">
                        <div class="form-card">
                            <div class="form-header">
                                <h3>${this.currentEditId ? '编辑记录' : '新增记录'}</h3>
                                <button class="btn btn-secondary" id="backToListBtn">← 返回列表</button>
                            </div>
                            <form id="guaranteeForm" class="guarantee-form">
                                <div class="form-grid">
                                    <div class="form-group">
                                        <label>登记时间 <span class="required">*</span></label>
                                        <input type="text" id="registerTime" value="${this.getTodayDate()}" required>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>担保类型 <span class="required">*</span></label>
                                        <select id="guaranteeType" required>
                                            <option value="">请选择</option>
                                            <option value="预留及备用金担保">预留及备用金担保</option>
                                            <option value="工资提成担保">工资提成担保</option>
                                            <option value="其他">其他</option>
                                        </select>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>对接人</label>
                                        <input type="text" id="contactPerson" placeholder="输入对接人姓名">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>大区</label>
                                        <input type="text" id="region" placeholder="输入大区名称">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>担保方 <span class="required">*</span></label>
                                        <div style="position: relative;">
                                            <input type="text" id="guarantor" placeholder="输入姓名或省区搜索，支持多选" required autocomplete="off">
                                            <div id="guarantorDropdown" style="display: none; position: absolute; top: 100%; left: 0; right: 0; background: var(--bg-card); border: 1px solid var(--border-color); border-radius: 8px; max-height: 300px; overflow-y: auto; z-index: 1000; margin-top: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                                                <div style="padding: 8px; border-bottom: 1px solid var(--border-color); display: flex; align-items: center; gap: 8px;">
                                                    <input type="checkbox" id="selectAllGuarantors" style="cursor: pointer;">
                                                    <label for="selectAllGuarantors" style="cursor: pointer; margin: 0; font-size: 0.85rem; color: var(--text-secondary);">全选</label>
                                                </div>
                                                <div id="guarantorOptions" style="max-height: 250px; overflow-y: auto;"></div>
                                            </div>
                                        </div>
                                        <small style="color: var(--text-muted); font-size: 0.8rem; margin-top: 4px; display: block;">
                                            输入姓名或省区搜索，支持多选（用逗号分隔）
                                        </small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>省区</label>
                                        <input type="text" id="province">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>离职标识</label>
                                        <select id="resignationFlag">
                                            <option value="">请选择</option>
                                            <option value="是">是</option>
                                            <option value="否">否</option>
                                        </select>
                                    </div>
                                    
                                    <!-- 工资提成担保相关字段 -->
                                    <div class="form-group" id="salaryCommissionAmountGroup" style="display: none;">
                                        <label>工资提成担保金额 <span class="required" id="salaryCommissionAmountRequired" style="display: none;">*</span></label>
                                        <input type="number" id="salaryCommissionAmount" step="0.01" min="0">
                                    </div>
                                    
                                    <!-- 预留及备用金担保相关字段 -->
                                    <div class="form-group" id="reservedAmountGroup" style="display: none;">
                                        <label>预留金额 <span class="required" id="reservedAmountRequired" style="display: none;">*</span></label>
                                        <input type="number" id="reservedAmount" step="0.01" min="0" readonly>
                                        <small style="color: var(--text-muted); font-size: 0.8rem; margin-top: 4px; display: block;">
                                            系统根据导入的预留数据自动计算（多人时自动累加）
                                        </small>
                                        <div id="guarantorReserveDetails" style="margin-top: 8px; padding: 8px; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; display: none;">
                                            <div style="font-weight: 600; margin-bottom: 4px; color: var(--text-primary);">各担保人预留金额明细：</div>
                                            <div id="guarantorReserveList" style="color: var(--text-secondary);"></div>
                                        </div>
                                    </div>
                                    
                                    <div class="form-group" id="reserveFundAmountGroup" style="display: none;">
                                        <label>备用金金额 <span class="required" id="reserveFundAmountRequired" style="display: none;">*</span></label>
                                        <input type="number" id="reserveFundAmount" step="0.01" min="0" placeholder="请输入备用金金额">
                                        <small style="color: var(--text-muted); font-size: 0.8rem; margin-top: 4px; display: block;">
                                            省区备用金或大区备用金需要手动输入金额
                                        </small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>担保金额 <span class="required">*</span></label>
                                        <input type="number" id="guaranteeAmount" step="0.01" min="0" required>
                                    </div>
                                    
                                    <div class="form-group" id="remainingAmountGroup" style="display: none;">
                                        <label>剩余可担保金额 <span class="required" id="remainingAmountRequired" style="display: none;">*</span></label>
                                        <input type="number" id="remainingAmount" step="0.01" min="0" readonly>
                                        <small style="color: var(--text-muted); font-size: 0.8rem; margin-top: 4px; display: block;">
                                            系统自动计算：预留金额 - 担保金额 + 预留回款金额
                                        </small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>担保客户</label>
                                        <input type="text" id="guaranteedCustomer">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>订单号</label>
                                        <input type="text" id="orderNumber">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>预计回款时间</label>
                                        <input type="date" id="expectedPaymentTime">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>审批编号</label>
                                        <input type="text" id="approvalNumber">
                                    </div>
                                    
                                    <div class="form-group full-width">
                                        <label>备注</label>
                                        <textarea id="remarks" rows="3"></textarea>
                                    </div>
                                    
                                    <div class="form-group" id="salaryCommissionPaymentGroup">
                                        <label>工资提成回款 <span class="required" id="salaryCommissionPaymentRequired" style="display: none;">*</span></label>
                                        <input type="number" id="salaryCommissionPayment" step="0.01" min="0">
                                        <small style="color: var(--text-muted); font-size: 0.8rem; margin-top: 4px; display: none;" id="salaryCommissionPaymentHint">
                                            工资提成担保类型下，客户回款时此字段为必填
                                        </small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>预留回款金额</label>
                                        <input type="number" id="reservedPaymentAmount" step="0.01" min="0">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>未回款金额</label>
                                        <input type="number" id="unpaidAmount" step="0.01" min="0">
                                    </div>
                                    
                                    <div class="form-group">
                                        <label>回款时间</label>
                                        <input type="month" id="paymentTime" placeholder="选择年月，如：2025-10">
                                    </div>
                                </div>
                                
                                <div class="form-actions">
                                    <button type="submit" class="btn btn-primary">
                                        ${this.currentEditId ? '更新记录' : '添加记录'}
                                    </button>
                                    <button type="button" class="btn btn-secondary" id="resetFormBtn">重置</button>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        this.bindFormEvents(container);
        this.initDatePickers();
        
        // 担保类型变化时，动态设置必填项
        const guaranteeTypeSelect = container.querySelector('#guaranteeType');
        if (guaranteeTypeSelect) {
            guaranteeTypeSelect.addEventListener('change', () => {
                this.handleGuaranteeTypeChange();
            });
            
            // 初始化时，如果有担保类型，触发一次变化事件
            if (guaranteeTypeSelect.value) {
                this.handleGuaranteeTypeChange();
            }
        }
        
        // 回款时间变化时，如果是工资提成担保，显示工资提成回款的必填标识
        // 使用事件委托，监听 flatpickr 的变化
        container.addEventListener('change', (e) => {
            if (e.target.id === 'paymentTime' || (e.target.classList && e.target.classList.contains('flatpickr-input'))) {
                const guaranteeType = document.getElementById('guaranteeType')?.value;
                const salaryCommissionPaymentRequired = document.getElementById('salaryCommissionPaymentRequired');
                if (guaranteeType === '工资提成担保' && salaryCommissionPaymentRequired) {
                    const paymentTime = document.getElementById('paymentTime').value;
                    if (paymentTime) {
                        salaryCommissionPaymentRequired.style.display = 'inline';
                    } else {
                        salaryCommissionPaymentRequired.style.display = 'none';
                    }
                }
            }
        });
    }
    
    initDatePickers() {
        // 初始化预计回款时间选择器（日期）
        const expectedPaymentTimeInput = document.getElementById('expectedPaymentTime');
        if (expectedPaymentTimeInput && !this.datePickerInstances.expectedPaymentTime) {
            this.datePickerInstances.expectedPaymentTime = flatpickr(expectedPaymentTimeInput, {
                dateFormat: 'Y-m-d',
                locale: 'zh',
                allowInput: false,
                clickOpens: true,
            });
        }
        
        // 回款时间使用month类型输入框，不需要日期选择器
        // 移除回款时间的日期选择器初始化
    }
    
    destroyDatePickers() {
        // 销毁所有日期选择器实例
        Object.values(this.datePickerInstances).forEach(instance => {
            if (instance && instance.destroy) {
                instance.destroy();
            }
        });
        this.datePickerInstances = {};
    }
    
    renderTableBody() {
        let filteredData = this.data;
        
        // 应用筛选
        if (this.filterGuarantor || this.filterGuaranteedCustomer || this.filterCustomer) {
            filteredData = filteredData.filter(record => {
                const matchGuarantor = !this.filterGuarantor || 
                    (record.guarantor && record.guarantor.includes(this.filterGuarantor));
                const matchGuaranteedCustomer = !this.filterGuaranteedCustomer || 
                    (record.guaranteedCustomer && record.guaranteedCustomer.includes(this.filterGuaranteedCustomer));
                const matchCustomer = !this.filterCustomer || 
                    (record.guaranteedCustomer && record.guaranteedCustomer.includes(this.filterCustomer));
                return matchGuarantor && matchGuaranteedCustomer && matchCustomer;
            });
        }
        
        if (filteredData.length === 0) {
            return '<tr><td colspan="19" class="empty-table">暂无数据或没有匹配的筛选结果</td></tr>';
        }
        
        const sortedData = this.getSortedData(filteredData);
        
        return sortedData.map(record => `
            <tr data-id="${record.id}">
                <td style="width: 140px; min-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${this.formatDateTime(record.registerTime)}">${this.formatDateTime(record.registerTime)}</td>
                <td style="width: 100px; min-width: 100px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.guarantor || '')}">${this.escapeHtml(record.guarantor || '')}</td>
                <td style="width: 120px; min-width: 120px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.guaranteeType || '')}">${this.escapeHtml(record.guaranteeType || '')}</td>
                <td style="width: 100px; min-width: 100px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.contactPerson || '')}">${this.escapeHtml(record.contactPerson || '')}</td>
                <td style="width: 100px; min-width: 100px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.region || '')}">${this.escapeHtml(record.region || '')}</td>
                <td style="width: 80px; min-width: 80px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.province || '')}">${this.escapeHtml(record.province || '')}</td>
                <td style="width: 90px; min-width: 90px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.resignationFlag || '')}">${this.escapeHtml(record.resignationFlag || '')}</td>
                <td style="width: 140px; min-width: 140px; text-align: right;">${this.formatNumber(record.salaryCommissionAmount)}</td>
                <td style="width: 100px; min-width: 100px; text-align: right;">${this.formatNumber(record.reservedAmount)}</td>
                <td style="width: 100px; min-width: 100px; text-align: right;">${this.formatNumber(record.guaranteeAmount)}</td>
                <td style="width: 130px; min-width: 130px; text-align: right;">${this.formatNumber(record.remainingAmount)}</td>
                <td style="width: 120px; min-width: 120px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.guaranteedCustomer || '')}">${this.escapeHtml(record.guaranteedCustomer || '')}</td>
                <td style="width: 110px; min-width: 110px; max-width: 110px; word-break: break-all; overflow: hidden; text-overflow: ellipsis; font-size: 0.75rem;" title="${this.escapeHtml(record.orderNumber || '')}">${this.escapeHtml(record.orderNumber || '')}</td>
                <td style="width: 120px; min-width: 120px; max-width: 120px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${this.formatDate(record.expectedPaymentTime)}">${this.formatDate(record.expectedPaymentTime)}</td>
                <td style="width: 120px; min-width: 120px; overflow: hidden; text-overflow: ellipsis;" title="${this.escapeHtml(record.approvalNumber || '')}">${this.escapeHtml(record.approvalNumber || '')}</td>
                <td class="remarks-cell" style="width: 150px; min-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${this.escapeHtml(record.remarks || '')}">
                    ${this.escapeHtml(record.remarks || '')}
                </td>
                <td style="width: 120px; min-width: 120px; text-align: right;">${this.formatNumber(record.salaryCommissionPayment)}</td>
                <td style="width: 120px; min-width: 120px; text-align: right;">${this.formatNumber(record.reservedPaymentAmount)}</td>
                <td style="width: 100px; min-width: 100px; text-align: right;">${this.formatNumber(record.unpaidAmount)}</td>
                <td style="width: 120px; min-width: 120px; max-width: 120px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${this.formatPaymentTime(record.paymentTime)}">${this.formatPaymentTime(record.paymentTime)}</td>
                <td class="actions-col" style="width: 100px; min-width: 100px; text-align: center;">
                    <button class="btn-icon btn-edit" data-id="${record.id}" title="编辑">✏️</button>
                    <button class="btn-icon btn-delete" data-id="${record.id}" title="删除">🗑️</button>
                </td>
            </tr>
        `).join('');
    }
    
    getSortedData(data = null) {
        const sourceData = data || this.data;
        
        if (!this.sortColumn) {
            return [...sourceData];
        }
        
        const sorted = [...sourceData].sort((a, b) => {
            let aVal = a[this.sortColumn];
            let bVal = b[this.sortColumn];
            
            // 处理日期时间
            if (this.sortColumn === 'registerTime' || this.sortColumn === 'expectedPaymentTime' || this.sortColumn === 'paymentTime') {
                aVal = aVal ? new Date(aVal).getTime() : 0;
                bVal = bVal ? new Date(bVal).getTime() : 0;
            }
            
            // 处理数字
            if (typeof aVal === 'number' || !isNaN(parseFloat(aVal))) {
                aVal = parseFloat(aVal) || 0;
                bVal = parseFloat(bVal) || 0;
            }
            
            // 处理字符串
            if (typeof aVal === 'string') {
                aVal = aVal.toLowerCase();
                bVal = (bVal || '').toLowerCase();
            }
            
            if (aVal < bVal) return this.sortDirection === 'asc' ? -1 : 1;
            if (aVal > bVal) return this.sortDirection === 'asc' ? 1 : -1;
            return 0;
        });
        
        return sorted;
    }
    
    formatDateTime(value) {
        if (!value) return '';
        const date = new Date(value);
        if (isNaN(date.getTime())) return value;
        return date.toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });
    }
    
    formatDate(value) {
        if (!value) return '';
        // 如果已经是YYYY-MM-DD格式，直接返回
        if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
            return value;
        }
        
        // 检测 Excel 日期序列号（大于 1 的数字，可能是日期序列号）
        const numValue = typeof value === 'number' ? value : parseFloat(value);
        if (!isNaN(numValue) && numValue > 1 && numValue < 1000000) {
            // Excel 日期序列号从 1900-01-01 开始（但 Excel 错误地认为 1900 是闰年）
            // 所以需要减去 2 天来修正
            const excelEpoch = new Date(1899, 11, 30); // 1900-01-01 的前一天
            const date = new Date(excelEpoch.getTime() + (numValue - 1) * 24 * 60 * 60 * 1000);
            if (!isNaN(date.getTime())) {
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const day = String(date.getDate()).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
        }
        
        // 尝试解析日期
        const date = new Date(value);
        if (isNaN(date.getTime())) {
            // 如果解析失败，尝试处理其他格式
            if (typeof value === 'string') {
                // 尝试处理 YYYY/MM/DD 格式
                const parts = value.split(/[\/\-]/);
                if (parts.length === 3) {
                    const year = parts[0].padStart(4, '0');
                    const month = parts[1].padStart(2, '0');
                    const day = parts[2].padStart(2, '0');
                    return `${year}-${month}-${day}`;
                }
            }
            return value;
        }
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
    
    formatPaymentTime(value) {
        if (!value) return '';
        // 如果已经是YYYY-MM格式，直接返回
        if (typeof value === 'string' && /^\d{4}-\d{2}$/.test(value)) {
            return value;
        }
        
        // 如果是YYYY-MM-DD格式，提取年月
        if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
            return value.substring(0, 7); // 返回 YYYY-MM
        }
        
        // 尝试解析日期并提取年月
        const date = new Date(value);
        if (!isNaN(date.getTime())) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            return `${year}-${month}`;
        }
        
        // 尝试处理其他格式（如 YYYY/MM 或 YYYY/MM/DD）
        if (typeof value === 'string') {
            const parts = value.split(/[\/\-]/);
            if (parts.length >= 2) {
                const year = parts[0].padStart(4, '0');
                const month = parts[1].padStart(2, '0');
                return `${year}-${month}`;
            }
        }
        
        return value;
    }
    
    formatNumber(value) {
        if (value === null || value === undefined || value === '') return '';
        const num = parseFloat(value);
        if (isNaN(num)) return '';
        return num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    bindListEvents(container) {
        // 新增按钮
        const addNewBtn = container.querySelector('#addNewBtn');
        if (addNewBtn) {
            addNewBtn.addEventListener('click', () => {
                this.currentEditId = null;
                this.currentView = 'form';
                const pageContainer = document.getElementById('page-content');
                if (pageContainer) {
                    this.render(pageContainer);
                }
            });
        }
        
        // 导出数据
        const exportBtn = container.querySelector('#exportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', async () => {
                await this.exportData();
            });
        }
        
        // 导入预留数据
        const importReserveBtn = container.querySelector('#importReserveBtn');
        if (importReserveBtn) {
            importReserveBtn.addEventListener('click', () => {
                this.importReserveData();
            });
        }
        
        // 导入历史数据
        const importHistoryBtn = container.querySelector('#importHistoryBtn');
        if (importHistoryBtn) {
            importHistoryBtn.addEventListener('click', () => {
                this.importHistoryData();
            });
        }
        
        // 筛选功能
        const applyFilterBtn = container.querySelector('#applyFilterBtn');
        const clearFilterBtn = container.querySelector('#clearFilterBtn');
        if (applyFilterBtn) {
            applyFilterBtn.addEventListener('click', () => {
                this.applyFilter();
            });
        }
        if (clearFilterBtn) {
            clearFilterBtn.addEventListener('click', () => {
                this.clearFilter();
            });
        }
        
        // 表格操作按钮
        const tableBody = container.querySelector('#guaranteeTableBody');
        if (tableBody) {
            tableBody.addEventListener('click', (e) => {
                const btn = e.target.closest('.btn-icon');
                if (!btn) return;
                
                const id = btn.dataset.id;
                if (!id) return;
                
                // 统一处理id，转换为数字或保持原样
                const recordId = typeof id === 'string' && !isNaN(id) ? parseFloat(id) : id;
                
                if (btn.classList.contains('btn-edit')) {
                    e.preventDefault();
                    e.stopPropagation();
                    this.editRecord(recordId);
                } else if (btn.classList.contains('btn-delete')) {
                    e.preventDefault();
                    e.stopPropagation();
                    this.confirmDelete(recordId);
                }
            });
        }
        
        // 表格排序
        const sortableHeaders = container.querySelectorAll('.sortable');
        sortableHeaders.forEach(header => {
            header.addEventListener('click', () => {
                const column = header.dataset.column;
                this.handleSort(column, header);
            });
        });
        
        // 列宽调整功能
        this.initColumnResizer(container);
    }
    
    bindFormEvents(container) {
        // 返回列表按钮
        const backToListBtn = container.querySelector('#backToListBtn');
        if (backToListBtn) {
            backToListBtn.addEventListener('click', () => {
                this.backToList();
            });
        }
        
        // 表单提交
        const form = container.querySelector('#guaranteeForm');
        if (form) {
            form.addEventListener('submit', (e) => {
                e.preventDefault();
                this.handleSubmit();
            });
        }
        
        // 重置表单
        const resetBtn = container.querySelector('#resetFormBtn');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                this.resetForm();
            });
        }
        
        // 担保金额变化时，自动计算剩余可担保金额
        const guaranteeAmountInput = container.querySelector('#guaranteeAmount');
        if (guaranteeAmountInput) {
            guaranteeAmountInput.addEventListener('input', () => {
                this.calculateRemainingAmount();
            });
            guaranteeAmountInput.addEventListener('blur', () => {
                this.calculateRemainingAmount();
            });
        }
        
        // 预留金额变化时，自动计算剩余可担保金额
        const reservedAmountInput = container.querySelector('#reservedAmount');
        if (reservedAmountInput) {
            reservedAmountInput.addEventListener('input', () => {
                this.calculateRemainingAmount();
            });
        }
        
        // 备用金金额变化时，自动计算剩余可担保金额
        const reserveFundAmountInput = container.querySelector('#reserveFundAmount');
        if (reserveFundAmountInput) {
            reserveFundAmountInput.addEventListener('input', () => {
                this.calculateRemainingAmount();
            });
        }
        
        // 预留回款金额变化时，自动计算剩余可担保金额
        const reservedPaymentAmountInput = container.querySelector('#reservedPaymentAmount');
        if (reservedPaymentAmountInput) {
            reservedPaymentAmountInput.addEventListener('input', () => {
                this.calculateRemainingAmount();
            });
        }
        
        // 初始化担保方下拉选择框
        this.initGuarantorDropdown(container);
        
        // 担保方变化时，自动填充预留金额并更新显示
        const guarantorInput = container.querySelector('#guarantor');
        if (guarantorInput) {
            guarantorInput.addEventListener('blur', () => {
                // 延迟隐藏下拉框，以便点击选项
                setTimeout(() => {
                    const dropdown = document.getElementById('guarantorDropdown');
                    if (dropdown) {
                        dropdown.style.display = 'none';
                    }
                }, 200);
                this.updateReserveAmountVisibility();
                this.autoFillReservedAmount();
            });
            guarantorInput.addEventListener('input', () => {
                // 实时更新显示状态
                const guaranteeType = document.getElementById('guaranteeType')?.value;
                if (guaranteeType === '预留及备用金担保') {
                    this.updateReserveAmountVisibility();
                    // 延迟一下再自动填充，避免输入过程中频繁触发
                    clearTimeout(this.autoFillTimeout);
                    this.autoFillTimeout = setTimeout(() => {
                        this.autoFillReservedAmount();
                    }, 500);
                }
            });
        }
    }
    
    applyFilter() {
        const filterGuarantorInput = document.getElementById('filterGuarantor');
        const filterGuaranteedCustomerInput = document.getElementById('filterGuaranteedCustomer');
        const filterCustomerInput = document.getElementById('filterCustomer');
        
        this.filterGuarantor = filterGuarantorInput ? filterGuarantorInput.value.trim() : '';
        this.filterGuaranteedCustomer = filterGuaranteedCustomerInput ? filterGuaranteedCustomerInput.value.trim() : '';
        this.filterCustomer = filterCustomerInput ? filterCustomerInput.value.trim() : '';
        
        this.refreshTable();
    }
    
    clearFilter() {
        const filterGuarantorInput = document.getElementById('filterGuarantor');
        const filterGuaranteedCustomerInput = document.getElementById('filterGuaranteedCustomer');
        const filterCustomerInput = document.getElementById('filterCustomer');
        
        if (filterGuarantorInput) filterGuarantorInput.value = '';
        if (filterGuaranteedCustomerInput) filterGuaranteedCustomerInput.value = '';
        if (filterCustomerInput) filterCustomerInput.value = '';
        
        this.filterGuarantor = '';
        this.filterGuaranteedCustomer = '';
        this.filterCustomer = '';
        
        this.refreshTable();
    }
    
    handleSubmit() {
        const form = document.getElementById('guaranteeForm');
        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }
        
        // 获取表单值（登记时间直接取原始值，不做格式转换）
        const registerTime = document.getElementById('registerTime').value.trim();
        
        let expectedPaymentTime = '';
        const expectedPaymentTimeInput = document.getElementById('expectedPaymentTime');
        if (this.datePickerInstances.expectedPaymentTime) {
            const selectedDate = this.datePickerInstances.expectedPaymentTime.selectedDates[0];
            if (selectedDate) {
                expectedPaymentTime = selectedDate.toISOString().split('T')[0];
            } else if (expectedPaymentTimeInput) {
                expectedPaymentTime = expectedPaymentTimeInput.value;
            }
        } else if (expectedPaymentTimeInput) {
            expectedPaymentTime = expectedPaymentTimeInput.value;
        }
        
        let paymentTime = '';
        const paymentTimeInput = document.getElementById('paymentTime');
        if (paymentTimeInput && paymentTimeInput.value) {
            // 回款时间使用年月格式（YYYY-MM）
            paymentTime = this.formatPaymentTime(paymentTimeInput.value);
        }
        
        const guaranteeType = document.getElementById('guaranteeType').value;
        
        // 如果是预留及备用金担保，验证担保方和必填项
        if (guaranteeType === '预留及备用金担保') {
            const guarantor = document.getElementById('guarantor').value.trim();
            if (!guarantor) {
                alert('担保方为必填项');
                return;
            }
            
            // 检查是否是备用金（省区备用金或大区备用金）
            const isReserveFund = guarantor.includes('省区备用金') || guarantor.includes('大区备用金');
            
            if (!isReserveFund) {
                // 不是备用金，需要验证是否在预留数据名单中
                const guarantorNames = guarantor.split(/[,，]/).map(name => name.trim()).filter(name => name);
                const invalidNames = [];
                
                for (const name of guarantorNames) {
                    const found = this.reserveData.find(r => 
                        r.guarantor && r.guarantor.trim() === name
                    );
                    if (!found) {
                        invalidNames.push(name);
                    }
                }
                
                if (invalidNames.length > 0) {
                    alert(`以下担保方不在预留数据名单中：${invalidNames.join('、')}\n请先导入预留数据或使用"省区备用金"、"大区备用金"`);
                    return;
                }
                
                // 不是备用金，需要验证预留金额和剩余可担保金额
                const reservedAmount = parseFloat(document.getElementById('reservedAmount').value) || 0;
                const guaranteeAmount = parseFloat(document.getElementById('guaranteeAmount').value) || 0;
                const remainingAmount = parseFloat(document.getElementById('remainingAmount').value) || 0;
                
                if (!reservedAmount || reservedAmount <= 0) {
                    alert('预留及备用金担保类型下，预留金额为必填项且必须大于0');
                    return;
                }
                if (!guaranteeAmount || guaranteeAmount <= 0) {
                    alert('担保金额为必填项且必须大于0');
                    return;
                }
                if (remainingAmount < 0) {
                    alert('剩余可担保金额不能为负数');
                    return;
                }
            } else {
                // 是备用金，需要验证备用金金额和担保金额
                const reserveFundAmount = parseFloat(document.getElementById('reserveFundAmount')?.value) || 0;
                const guaranteeAmount = parseFloat(document.getElementById('guaranteeAmount').value) || 0;
                
                if (!reserveFundAmount || reserveFundAmount <= 0) {
                    alert('备用金金额为必填项且必须大于0');
                    return;
                }
                if (!guaranteeAmount || guaranteeAmount <= 0) {
                    alert('担保金额为必填项且必须大于0');
                    return;
                }
            }
        }
        
        // 如果是工资提成担保，验证必填项并检查未回款记录
        if (guaranteeType === '工资提成担保') {
            const salaryCommissionAmount = parseFloat(document.getElementById('salaryCommissionAmount').value) || 0;
            
            if (!salaryCommissionAmount || salaryCommissionAmount <= 0) {
                alert('工资提成担保类型下，工资提成担保金额为必填项且必须大于0');
                return;
            }
            
            // 检查是否有未回款的工资提成担保
            const guarantor = document.getElementById('guarantor').value.trim();
            if (guarantor) {
                const guarantorNames = guarantor.split(/[,，]/).map(name => name.trim()).filter(name => name);
                const unpaidRecords = [];
                
                for (const name of guarantorNames) {
                    const records = this.data.filter(r => 
                        r.guaranteeType === '工资提成担保' &&
                        r.guarantor && r.guarantor.includes(name) &&
                        (!r.paymentTime || r.paymentTime === '') &&
                        r.salaryCommissionAmount > 0
                    );
                    
                    if (records.length > 0) {
                        const totalUnpaid = records.reduce((sum, r) => sum + (parseFloat(r.salaryCommissionAmount) || 0), 0);
                        unpaidRecords.push({
                            name: name,
                            amount: totalUnpaid,
                            count: records.length
                        });
                    }
                }
                
                if (unpaidRecords.length > 0) {
                    const message = unpaidRecords.map(r => 
                        `${r.name}：存在${r.count}笔未回款工资提成担保，金额${r.amount.toFixed(2)}元`
                    ).join('\n');
                    if (!confirm(`⚠️ 此员工存在未回款工资提成担保：\n\n${message}\n\n是否继续添加？`)) {
                        return;
                    }
                }
            }
            
            // 如果已填写回款时间，则工资提成回款为必填
            const paymentTime = document.getElementById('paymentTime').value;
            if (paymentTime) {
                const salaryCommissionPayment = parseFloat(document.getElementById('salaryCommissionPayment').value) || 0;
                if (!salaryCommissionPayment || salaryCommissionPayment <= 0) {
                    alert('工资提成担保类型下，客户回款时工资提成回款为必填项');
                    return;
                }
            }
        }
        
        const formData = {
            registerTime: registerTime,
            guarantor: document.getElementById('guarantor').value.trim(),
            guaranteeType: guaranteeType,
            contactPerson: document.getElementById('contactPerson')?.value.trim() || '',
            region: document.getElementById('region')?.value.trim() || '',
            province: document.getElementById('province').value.trim(),
            resignationFlag: document.getElementById('resignationFlag').value,
            useSalaryGuarantee: document.getElementById('useSalaryGuarantee').value,
            useCommissionGuarantee: document.getElementById('useCommissionGuarantee').value,
            salaryCommissionAmount: parseFloat(document.getElementById('salaryCommissionAmount').value) || 0,
            reservedAmount: parseFloat(document.getElementById('reservedAmount').value) || 0,
            reserveFundAmount: parseFloat(document.getElementById('reserveFundAmount')?.value) || 0,
            guaranteeAmount: parseFloat(document.getElementById('guaranteeAmount').value) || 0,
            remainingAmount: parseFloat(document.getElementById('remainingAmount').value) || 0,
            guaranteedCustomer: document.getElementById('guaranteedCustomer').value.trim(),
            orderNumber: document.getElementById('orderNumber').value.trim(),
            expectedPaymentTime: expectedPaymentTime,
            approvalNumber: document.getElementById('approvalNumber').value.trim(),
            remarks: document.getElementById('remarks').value.trim(),
            salaryCommissionPayment: parseFloat(document.getElementById('salaryCommissionPayment').value) || 0,
            reservedPaymentAmount: parseFloat(document.getElementById('reservedPaymentAmount').value) || 0,
            unpaidAmount: parseFloat(document.getElementById('unpaidAmount').value) || 0,
            paymentTime: paymentTime
        };
        
        if (this.currentEditId) {
            // 更新记录
            const index = this.data.findIndex(r => r.id === this.currentEditId);
            if (index !== -1) {
                this.data[index] = { ...this.data[index], ...formData };
                this.saveData();
                this.showToast('✅ 记录已更新');
                // 返回列表视图
                this.currentEditId = null;
                this.currentView = 'list';
                const container = document.getElementById('page-content');
                if (container) {
                    this.render(container);
                }
            }
        } else {
            // 新增记录
            const newRecord = {
                id: Date.now() + Math.random(),
                ...formData
            };
            this.data.push(newRecord);
            this.saveData();
            this.showToast('✅ 记录已添加');
            // 返回列表视图
            this.currentView = 'list';
            const container = document.getElementById('page-content');
            if (container) {
                this.render(container);
            }
        }
    }
    
    editRecord(id) {
        // 统一处理id类型
        const recordId = typeof id === 'string' && !isNaN(id) ? parseFloat(id) : id;
        const record = this.data.find(r => {
            // 支持数字和字符串类型的id匹配
            return r.id === recordId || r.id === id || String(r.id) === String(id);
        });
        if (!record) {
            console.error('未找到记录，id:', id, 'recordId:', recordId);
            return;
        }
        
        this.currentEditId = parseInt(id);
        this.currentView = 'form';
        const container = document.getElementById('page-content');
        if (container) {
            // 先销毁旧的日期选择器
            this.destroyDatePickers();
            this.render(container);
        }
        
        // 等待DOM更新后再填充表单
        setTimeout(() => {
            // 填充表单（登记时间直接使用原始值，不做格式转换）
            const registerTimeInput = document.getElementById('registerTime');
            if (registerTimeInput) {
                registerTimeInput.value = record.registerTime || '';
            }
            
            const guarantorInput = document.getElementById('guarantor');
            if (guarantorInput) {
                guarantorInput.value = record.guarantor || '';
            }
            
            const guaranteeTypeSelect = document.getElementById('guaranteeType');
            if (guaranteeTypeSelect) {
                guaranteeTypeSelect.value = record.guaranteeType || '';
                // 触发担保类型变化事件，更新必填项显示
                this.handleGuaranteeTypeChange();
            }
            
            const contactPersonInput = document.getElementById('contactPerson');
            if (contactPersonInput) {
                contactPersonInput.value = record.contactPerson || '';
            }
            
            const regionInput = document.getElementById('region');
            if (regionInput) {
                regionInput.value = record.region || '';
            }
            
            const provinceInput = document.getElementById('province');
            if (provinceInput) {
                provinceInput.value = record.province || '';
            }
            
            const resignationFlagSelect = document.getElementById('resignationFlag');
            if (resignationFlagSelect) {
                resignationFlagSelect.value = record.resignationFlag || '';
            }
            
            const salaryCommissionAmountInput = document.getElementById('salaryCommissionAmount');
            if (salaryCommissionAmountInput) {
                salaryCommissionAmountInput.value = record.salaryCommissionAmount || '';
            }
            
            const reservedAmountInput = document.getElementById('reservedAmount');
            if (reservedAmountInput) {
                reservedAmountInput.value = record.reservedAmount || '';
            }
            
            const reserveFundAmountInput = document.getElementById('reserveFundAmount');
            if (reserveFundAmountInput) {
                reserveFundAmountInput.value = record.reserveFundAmount || '';
            }
            
            const guaranteeAmountInput = document.getElementById('guaranteeAmount');
            if (guaranteeAmountInput) {
                guaranteeAmountInput.value = record.guaranteeAmount || '';
            }
            
            const remainingAmountInput = document.getElementById('remainingAmount');
            if (remainingAmountInput) {
                remainingAmountInput.value = record.remainingAmount || '';
            }
            
            const guaranteedCustomerInput = document.getElementById('guaranteedCustomer');
            if (guaranteedCustomerInput) {
                guaranteedCustomerInput.value = record.guaranteedCustomer || '';
            }
            
            const orderNumberInput = document.getElementById('orderNumber');
            if (orderNumberInput) {
                orderNumberInput.value = record.orderNumber || '';
            }
            
            if (this.datePickerInstances.expectedPaymentTime && record.expectedPaymentTime) {
                this.datePickerInstances.expectedPaymentTime.setDate(record.expectedPaymentTime, false);
            } else {
                const expectedPaymentTimeInput = document.getElementById('expectedPaymentTime');
                if (expectedPaymentTimeInput) {
                    expectedPaymentTimeInput.value = record.expectedPaymentTime || '';
                }
            }
            
            const approvalNumberInput = document.getElementById('approvalNumber');
            if (approvalNumberInput) {
                approvalNumberInput.value = record.approvalNumber || '';
            }
            
            const remarksTextarea = document.getElementById('remarks');
            if (remarksTextarea) {
                remarksTextarea.value = record.remarks || '';
            }
            
            const salaryCommissionPaymentInput = document.getElementById('salaryCommissionPayment');
            if (salaryCommissionPaymentInput) {
                salaryCommissionPaymentInput.value = record.salaryCommissionPayment || '';
            }
            
            const reservedPaymentAmountInput = document.getElementById('reservedPaymentAmount');
            if (reservedPaymentAmountInput) {
                reservedPaymentAmountInput.value = record.reservedPaymentAmount || '';
            }
            
            const unpaidAmountInput = document.getElementById('unpaidAmount');
            if (unpaidAmountInput) {
                unpaidAmountInput.value = record.unpaidAmount || '';
            }
            
            const paymentTimeInput = document.getElementById('paymentTime');
            if (paymentTimeInput && record.paymentTime) {
                // 回款时间使用年月格式，如果是完整日期则提取年月
                paymentTimeInput.value = this.formatPaymentTime(record.paymentTime);
            }
            
            // 滚动到表单
            const formSection = document.querySelector('.guarantee-form-section');
            if (formSection) {
                formSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }, 0);
    }
    
    confirmDelete(id) {
        // 统一处理id类型
        const recordId = typeof id === 'string' && !isNaN(id) ? parseFloat(id) : id;
        const record = this.data.find(r => {
            const rId = r.id;
            return rId === recordId || rId === id || String(rId) === String(id) || String(rId) === String(recordId);
        });
        
        if (!record) {
            this.showError('未找到要删除的记录');
            return;
        }
        
        // 创建确认对话框
        const dialog = document.createElement('div');
        dialog.className = 'delete-confirm-dialog';
        dialog.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 10000;
        `;
        
        const dialogContent = document.createElement('div');
        dialogContent.className = 'delete-confirm-content';
        dialogContent.style.cssText = `
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            max-width: 500px;
            width: 90%;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        `;
        
        const recordInfo = `
            <div style="margin-bottom: 16px;">
                <strong>担保方：</strong>${this.escapeHtml(record.guarantor || '')}<br>
                <strong>担保客户：</strong>${this.escapeHtml(record.guaranteedCustomer || '')}<br>
                <strong>担保类型：</strong>${this.escapeHtml(record.guaranteeType || '')}<br>
                <strong>登记时间：</strong>${this.formatDateTime(record.registerTime)}
            </div>
        `;
        
        dialogContent.innerHTML = `
            <div style="margin-bottom: 20px;">
                <h3 style="margin: 0 0 16px 0; color: var(--accent-rose); font-size: 1.2rem;">⚠️ 确认删除</h3>
                <p style="margin: 0 0 12px 0; color: var(--text-primary);">确定要删除以下记录吗？此操作不可恢复！</p>
                ${recordInfo}
            </div>
            <div style="display: flex; gap: 12px; justify-content: flex-end;">
                <button class="btn btn-secondary" id="cancelDeleteBtn" style="padding: 8px 16px;">取消</button>
                <button class="btn btn-danger" id="confirmDeleteBtn" style="padding: 8px 16px; background: var(--accent-rose);">确认删除</button>
            </div>
        `;
        
        dialog.appendChild(dialogContent);
        document.body.appendChild(dialog);
        
        // 绑定事件
        const cancelBtn = dialog.querySelector('#cancelDeleteBtn');
        const confirmBtn = dialog.querySelector('#confirmDeleteBtn');
        
        const closeDialog = () => {
            document.body.removeChild(dialog);
        };
        
        cancelBtn.addEventListener('click', closeDialog);
        confirmBtn.addEventListener('click', () => {
            closeDialog();
            this.deleteRecord(id);
        });
        
        // 点击背景关闭
        dialog.addEventListener('click', (e) => {
            if (e.target === dialog) {
                closeDialog();
            }
        });
    }
    
    deleteRecord(id) {
        // 统一处理id类型，支持数字和字符串类型的id匹配
        const recordId = typeof id === 'string' && !isNaN(id) ? parseFloat(id) : id;
        this.data = this.data.filter(r => {
            // 保留id不匹配的记录（删除匹配的记录）
            const rId = r.id;
            return rId !== recordId && rId !== id && String(rId) !== String(id) && String(rId) !== String(recordId);
        });
        this.saveData();
        this.showToast('✅ 记录已删除');
        if (this.currentView === 'list') {
            this.refreshTable();
        } else {
            // 如果在表单视图，切换到列表视图
            this.currentView = 'list';
            const pageContainer = document.getElementById('page-content');
            if (pageContainer) {
                this.render(pageContainer);
            }
        }
    }
    
    clearAll() {
        this.data = [];
        this.saveData();
        this.showToast('✅ 所有记录已清空');
        this.currentView = 'list';
        const container = document.getElementById('page-content');
        if (container) {
            this.render(container);
        }
    }
    
    backToList() {
        this.currentEditId = null;
        this.currentView = 'list';
        const container = document.getElementById('page-content');
        if (container) {
            this.render(container);
        }
    }
    
    cancelEdit() {
        this.backToList();
    }
    
    resetForm() {
        document.getElementById('guaranteeForm').reset();
        this.currentEditId = null;
        
        // 设置登记时间为当天
        const registerTimeInput = document.getElementById('registerTime');
        if (registerTimeInput) {
            registerTimeInput.value = this.getTodayDate();
        }
        
        // 清空日期选择器
        if (this.datePickerInstances.expectedPaymentTime) {
            this.datePickerInstances.expectedPaymentTime.clear();
        }
        if (this.datePickerInstances.paymentTime) {
            this.datePickerInstances.paymentTime.clear();
        }
    }
    
    initGuarantorDropdown(container) {
        const guarantorInput = container.querySelector('#guarantor');
        const dropdown = document.getElementById('guarantorDropdown');
        const optionsContainer = document.getElementById('guarantorOptions');
        const selectAllCheckbox = document.getElementById('selectAllGuarantors');
        
        if (!guarantorInput || !dropdown || !optionsContainer) return;
        
        // 获取所有可用的担保方（从预留数据和现有数据中提取）
        const getAllGuarantors = () => {
            const guarantors = new Set();
            
            // 从预留数据中获取
            this.reserveData.forEach(item => {
                if (item.guarantor) {
                    guarantors.add(item.guarantor.trim());
                }
            });
            
            // 从现有数据中获取
            this.data.forEach(record => {
                if (record.guarantor) {
                    record.guarantor.split(/[,，]/).forEach(name => {
                        const trimmed = name.trim();
                        if (trimmed) {
                            guarantors.add(trimmed);
                        }
                    });
                }
            });
            
            // 添加备用金选项
            guarantors.add('省区备用金');
            guarantors.add('大区备用金');
            
            return Array.from(guarantors).sort();
        };
        
        // 获取所有省区（从现有数据中提取）
        const getAllProvinces = () => {
            const provinces = new Set();
            this.data.forEach(record => {
                if (record.province) {
                    provinces.add(record.province.trim());
                }
            });
            return Array.from(provinces).sort();
        };
        
        // 根据省区获取该省区的所有担保方
        const getGuarantorsByProvince = (province) => {
            const guarantors = new Set();
            this.data.forEach(record => {
                if (record.province === province && record.guarantor) {
                    record.guarantor.split(/[,，]/).forEach(name => {
                        const trimmed = name.trim();
                        if (trimmed) {
                            guarantors.add(trimmed);
                        }
                    });
                }
            });
            return Array.from(guarantors).sort();
        };
        
        // 渲染选项列表
        const renderOptions = (searchText = '') => {
            const allGuarantors = getAllGuarantors();
            const allProvinces = getAllProvinces();
            let filteredGuarantors = [];
            
            if (!searchText) {
                filteredGuarantors = allGuarantors;
            } else {
                const searchLower = searchText.toLowerCase();
                
                // 检查是否是省区名称
                const matchedProvince = allProvinces.find(p => p.toLowerCase().includes(searchLower));
                if (matchedProvince) {
                    // 如果是省区，显示该省区的所有担保方
                    filteredGuarantors = getGuarantorsByProvince(matchedProvince);
                } else {
                    // 否则按姓名搜索
                    filteredGuarantors = allGuarantors.filter(name => 
                        name.toLowerCase().includes(searchLower)
                    );
                }
            }
            
            // 获取当前已选中的担保方
            const selectedValues = guarantorInput.value.split(/[,，]/).map(v => v.trim()).filter(v => v);
            
            // 渲染选项
            optionsContainer.innerHTML = filteredGuarantors.map(name => {
                const isSelected = selectedValues.includes(name);
                const bgStyle = isSelected ? 'background: var(--bg-card-hover);' : '';
                return `
                    <div style="padding: 8px 12px; cursor: pointer; display: flex; align-items: center; gap: 8px; border-bottom: 1px solid var(--border-color); transition: background var(--transition-fast); ${bgStyle}" 
                         class="guarantor-option" data-value="${this.escapeHtml(name)}">
                        <input type="checkbox" ${isSelected ? 'checked' : ''} style="cursor: pointer;">
                        <span style="flex: 1; font-size: 0.9rem;">${this.escapeHtml(name)}</span>
                    </div>
                `;
            }).join('');
            
            // 更新全选状态
            if (selectAllCheckbox) {
                const allChecked = filteredGuarantors.length > 0 && 
                    filteredGuarantors.every(name => selectedValues.includes(name));
                selectAllCheckbox.checked = allChecked;
            }
            
            // 绑定选项点击事件
            optionsContainer.querySelectorAll('.guarantor-option').forEach(option => {
                option.addEventListener('click', (e) => {
                    if (e.target.tagName === 'INPUT') return;
                    
                    const checkbox = option.querySelector('input[type="checkbox"]');
                    const value = option.dataset.value;
                    
                    checkbox.checked = !checkbox.checked;
                    updateSelectedValues();
                });
                
                const checkbox = option.querySelector('input[type="checkbox"]');
                checkbox.addEventListener('change', () => {
                    updateSelectedValues();
                });
            });
        };
        
        // 更新选中的值
        const updateSelectedValues = () => {
            const selected = [];
            optionsContainer.querySelectorAll('.guarantor-option input[type="checkbox"]:checked').forEach(checkbox => {
                const option = checkbox.closest('.guarantor-option');
                if (option) {
                    selected.push(option.dataset.value);
                }
            });
            guarantorInput.value = selected.join('，');
            
            // 更新全选状态
            const allOptions = optionsContainer.querySelectorAll('.guarantor-option');
            const allChecked = allOptions.length > 0 && 
                Array.from(allOptions).every(opt => opt.querySelector('input[type="checkbox"]').checked);
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = allChecked;
            }
            
            // 触发自动填充预留金额
            this.updateReserveAmountVisibility();
            this.autoFillReservedAmount();
        };
        
        // 全选/取消全选
        if (selectAllCheckbox) {
            selectAllCheckbox.addEventListener('change', () => {
                const isChecked = selectAllCheckbox.checked;
                optionsContainer.querySelectorAll('.guarantor-option input[type="checkbox"]').forEach(checkbox => {
                    checkbox.checked = isChecked;
                });
                updateSelectedValues();
            });
        }
        
        // 输入时显示下拉框并搜索
        guarantorInput.addEventListener('input', (e) => {
            const searchText = e.target.value;
            renderOptions(searchText);
            dropdown.style.display = 'block';
        });
        
        // 聚焦时显示下拉框
        guarantorInput.addEventListener('focus', () => {
            renderOptions(guarantorInput.value);
            dropdown.style.display = 'block';
        });
        
        // 点击外部时隐藏下拉框
        document.addEventListener('click', (e) => {
            if (!guarantorInput.contains(e.target) && !dropdown.contains(e.target)) {
                dropdown.style.display = 'none';
            }
        });
        
        // 初始化渲染
        renderOptions();
    }
    
    refreshTable() {
        if (this.currentView !== 'list') return;
        
        const container = document.getElementById('page-content');
        if (!container) return;
        
        const tableBody = container.querySelector('#guaranteeTableBody');
        const recordCount = container.querySelector('.record-count');
        
        if (tableBody) {
            tableBody.innerHTML = this.renderTableBody();
        }
        
        if (recordCount) {
            recordCount.textContent = `共 ${this.data.length} 条记录`;
        }
        
        if (recordCount) {
            recordCount.textContent = `共 ${this.data.length} 条记录`;
        }
        
        // 更新按钮状态
        const exportBtn = document.getElementById('exportBtn');
        if (exportBtn) {
            exportBtn.disabled = this.data.length === 0;
        }
    }
    
    initColumnResizer(container) {
        const table = container.querySelector('.guarantee-table');
        if (!table) return;
        
        const headers = table.querySelectorAll('thead th');
        let isResizing = false;
        let currentHeader = null;
        let startX = 0;
        let startWidth = 0;
        let columnIndex = 0;
        
        headers.forEach((header, index) => {
            const resizer = header.querySelector('.resizer');
            if (!resizer) return;
            
            resizer.addEventListener('mousedown', (e) => {
                e.preventDefault();
                e.stopPropagation();
                
                isResizing = true;
                currentHeader = header;
                columnIndex = index;
                startX = e.pageX;
                startWidth = header.offsetWidth;
                
                header.classList.add('resizing');
                document.body.style.cursor = 'col-resize';
                document.body.style.userSelect = 'none';
            });
        });
        
        document.addEventListener('mousemove', (e) => {
            if (!isResizing || !currentHeader) return;
            
            const diff = e.pageX - startX;
            const newWidth = Math.max(50, startWidth + diff); // 最小宽度50px
            
            // 更新表头列宽
            currentHeader.style.width = newWidth + 'px';
            currentHeader.style.minWidth = newWidth + 'px';
            if (currentHeader.style.maxWidth) {
                currentHeader.style.maxWidth = newWidth + 'px';
            }
            
            // 同步更新表体对应列的宽度
            const tableBody = table.querySelector('tbody');
            if (tableBody) {
                const rows = tableBody.querySelectorAll('tr');
                rows.forEach(row => {
                    const cell = row.cells[columnIndex];
                    if (cell) {
                        cell.style.width = newWidth + 'px';
                        cell.style.minWidth = newWidth + 'px';
                        if (cell.style.maxWidth) {
                            cell.style.maxWidth = newWidth + 'px';
                        }
                    }
                });
            }
        });
        
        document.addEventListener('mouseup', () => {
            if (isResizing && currentHeader) {
                currentHeader.classList.remove('resizing');
                currentHeader = null;
            }
            isResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        });
    }
    
    handleSort(column, header) {
        // 更新排序状态
        if (this.sortColumn === column) {
            this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortColumn = column;
            this.sortDirection = 'asc';
        }
        
        // 更新所有排序图标的显示
        document.querySelectorAll('.sortable').forEach(h => {
            const icon = h.querySelector('.sort-icon');
            if (h === header) {
                icon.textContent = this.sortDirection === 'asc' ? '▲' : '▼';
                h.classList.add('active');
            } else {
                icon.textContent = '▼';
                h.classList.remove('active');
            }
        });
        
        // 刷新表格
        this.refreshTable();
    }
    
    async exportData() {
        if (this.data.length === 0) {
            this.showToast('暂无数据可导出');
            return;
        }
        
        // 检查XLSX库是否加载
        if (!window.XLSX) {
            this.showError('XLSX 库未加载，无法导出 Excel 文件');
            return;
        }
        
        try {
            // 准备表头
            const headers = [
                '登记时间', '担保方', '担保类型', '对接人', '大区', '省区', '离职标识',
                '工资提成担保金额', '预留金额', '担保金额', '剩余可担保金额', '担保客户', '订单号',
                '预计回款时间', '审批编号', '备注', '工资提成回款', '预留回款金额', '未回款金额', '回款时间'
            ];
            
            // 准备数据行
            const rows = this.data.map(record => [
                this.formatDateTime(record.registerTime),
                record.guarantor || '',
                record.guaranteeType || '',
                record.contactPerson || '',
                record.region || '',
                record.province || '',
                record.resignationFlag || '',
                record.salaryCommissionAmount || 0,
                record.reservedAmount || 0,
                record.guaranteeAmount || 0,
                record.remainingAmount || 0,
                record.guaranteedCustomer || '',
                record.orderNumber || '',
                this.formatDate(record.expectedPaymentTime),
                record.approvalNumber || '',
                record.remarks || '',
                record.salaryCommissionPayment || 0,
                record.reservedPaymentAmount || 0,
                record.unpaidAmount || 0,
                this.formatPaymentTime(record.paymentTime)
            ]);
            
            // 创建工作簿
            const wb = XLSX.utils.book_new();
            
            // 创建工作表数据（包含表头）
            const wsData = [headers, ...rows];
            
            // 创建工作表
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            
            // 设置列宽（可选，让Excel自动调整）
            const colWidths = headers.map(() => ({ wch: 15 }));
            ws['!cols'] = colWidths;
            
            // 将工作表添加到工作簿
            XLSX.utils.book_append_sheet(wb, ws, '担保台账');
            
            // 生成Excel文件
            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            
            // 检查是否在 Tauri 环境中
            if (window.__TAURI__) {
                const { invoke } = window.__TAURI__.core;
                const { save } = window.__TAURI__.dialog;
                
                // 打开保存文件对话框
                const filePath = await save({
                    defaultPath: `担保台账_${new Date().toISOString().split('T')[0]}.xlsx`,
                    filters: [{
                        name: 'Excel文件',
                        extensions: ['xlsx']
                    }]
                });
                
                if (filePath) {
                    // 将 ArrayBuffer 转换为 base64
                    const base64 = this.arrayBufferToBase64(excelBuffer);
                    
                    // 调用后端命令保存文件
                    await invoke('save_excel_file', {
                        filePath: filePath,
                        content: base64
                    });
                    
                    this.showToast('✅ 数据已导出为Excel文件');
                } else {
                    // 用户取消了保存操作
                    this.showToast('已取消导出');
                }
            } else {
                // 非 Tauri 环境，使用浏览器下载
                const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `担保台账_${new Date().toISOString().split('T')[0]}.xlsx`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                this.showToast('✅ 数据已导出为Excel文件');
            }
        } catch (error) {
            console.error('导出Excel失败:', error);
            if (error !== '用户取消操作') {
                this.showError('导出Excel文件失败: ' + error.message);
            }
        }
    }
    
    /**
     * 将 ArrayBuffer 转换为 Base64
     */
    arrayBufferToBase64(buffer) {
        let binary = '';
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return window.btoa(binary);
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
            border: 1px solid var(--border-color);
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
        }, 2500);
    }
    
    showError(message) {
        alert(message);
    }
    
    handleGuaranteeTypeChange() {
        const guaranteeType = document.getElementById('guaranteeType').value;
        
        // 预留及备用金担保相关字段
        const reservedAmountGroup = document.getElementById('reservedAmountGroup');
        const reserveFundAmountGroup = document.getElementById('reserveFundAmountGroup');
        const remainingAmountGroup = document.getElementById('remainingAmountGroup');
        const reservedAmountRequired = document.getElementById('reservedAmountRequired');
        const reserveFundAmountRequired = document.getElementById('reserveFundAmountRequired');
        const remainingAmountRequired = document.getElementById('remainingAmountRequired');
        const reservedAmountInput = document.getElementById('reservedAmount');
        const remainingAmountInput = document.getElementById('remainingAmount');
        
        // 工资提成担保相关字段
        const useSalaryGuaranteeGroup = document.getElementById('useSalaryGuaranteeGroup');
        const useCommissionGuaranteeGroup = document.getElementById('useCommissionGuaranteeGroup');
        const salaryCommissionAmountGroup = document.getElementById('salaryCommissionAmountGroup');
        const useSalaryGuaranteeRequired = document.getElementById('useSalaryGuaranteeRequired');
        const useCommissionGuaranteeRequired = document.getElementById('useCommissionGuaranteeRequired');
        const salaryCommissionAmountRequired = document.getElementById('salaryCommissionAmountRequired');
        const useSalaryGuaranteeSelect = document.getElementById('useSalaryGuarantee');
        const useCommissionGuaranteeSelect = document.getElementById('useCommissionGuarantee');
        const salaryCommissionAmountInput = document.getElementById('salaryCommissionAmount');
        
        // 工资提成回款字段
        const salaryCommissionPaymentRequired = document.getElementById('salaryCommissionPaymentRequired');
        const salaryCommissionPaymentHint = document.getElementById('salaryCommissionPaymentHint');
        
        if (guaranteeType === '预留及备用金担保') {
            // 隐藏工资提成担保金额字段
            if (salaryCommissionAmountGroup) salaryCommissionAmountGroup.style.display = 'none';
            if (salaryCommissionAmountInput) salaryCommissionAmountInput.value = '';
            
            // 隐藏工资提成回款的必填标识和提示
            if (salaryCommissionPaymentRequired) salaryCommissionPaymentRequired.style.display = 'none';
            if (salaryCommissionPaymentHint) salaryCommissionPaymentHint.style.display = 'none';
            
            // 默认显示预留金额和剩余可担保金额字段（除非是备用金）
            const guarantor = document.getElementById('guarantor')?.value.trim() || '';
            const isReserveFund = guarantor && (guarantor.includes('省区备用金') || guarantor.includes('大区备用金'));
            
            if (!isReserveFund) {
                // 不是备用金，显示预留金额和剩余可担保金额，隐藏备用金金额
                if (reservedAmountGroup) {
                    reservedAmountGroup.style.display = 'flex';
                }
                if (remainingAmountGroup) {
                    remainingAmountGroup.style.display = 'flex';
                }
                if (reservedAmountRequired) reservedAmountRequired.style.display = 'inline';
                if (remainingAmountRequired) remainingAmountRequired.style.display = 'inline';
                if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'none';
                
                // 自动填充预留金额
                if (guarantor) {
                    this.autoFillReservedAmount();
                }
            } else {
                // 是备用金，显示备用金金额和剩余可担保金额，隐藏预留金额
                if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'flex';
                if (reserveFundAmountRequired) reserveFundAmountRequired.style.display = 'inline';
                if (remainingAmountGroup) remainingAmountGroup.style.display = 'flex';
                if (remainingAmountRequired) remainingAmountRequired.style.display = 'inline';
                if (reservedAmountGroup) reservedAmountGroup.style.display = 'none';
            }
        } else if (guaranteeType === '工资提成担保') {
            // 隐藏预留及备用金担保相关字段
            if (reservedAmountGroup) reservedAmountGroup.style.display = 'none';
            if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'none';
            if (remainingAmountGroup) remainingAmountGroup.style.display = 'none';
            
            // 清空预留及备用金担保字段的值
            if (reservedAmountInput) reservedAmountInput.value = '';
            const reserveFundAmountInput = document.getElementById('reserveFundAmount');
            if (reserveFundAmountInput) reserveFundAmountInput.value = '';
            if (remainingAmountInput) remainingAmountInput.value = '';
            
            // 显示工资提成担保相关字段
            if (salaryCommissionAmountGroup) salaryCommissionAmountGroup.style.display = 'flex';
            if (salaryCommissionAmountRequired) salaryCommissionAmountRequired.style.display = 'inline';
            
            // 显示工资提成回款的提示
            if (salaryCommissionPaymentHint) salaryCommissionPaymentHint.style.display = 'block';
            
            // 如果已填写回款时间，显示工资提成回款的必填标识
            const paymentTime = document.getElementById('paymentTime').value;
            if (paymentTime && salaryCommissionPaymentRequired) {
                salaryCommissionPaymentRequired.style.display = 'inline';
            }
        } else {
            // 其他类型：显示所有字段，但不设置必填
            if (reservedAmountGroup) reservedAmountGroup.style.display = 'flex';
            if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'flex';
            if (remainingAmountGroup) remainingAmountGroup.style.display = 'flex';
            if (salaryCommissionAmountGroup) salaryCommissionAmountGroup.style.display = 'flex';
            
            // 隐藏所有必填标识
            if (reservedAmountRequired) reservedAmountRequired.style.display = 'none';
            if (reserveFundAmountRequired) reserveFundAmountRequired.style.display = 'none';
            if (remainingAmountRequired) remainingAmountRequired.style.display = 'none';
            if (salaryCommissionAmountRequired) salaryCommissionAmountRequired.style.display = 'none';
            if (salaryCommissionPaymentRequired) salaryCommissionPaymentRequired.style.display = 'none';
            if (salaryCommissionPaymentHint) salaryCommissionPaymentHint.style.display = 'none';
        }
    }
    
    updateReserveAmountVisibility() {
        const guarantor = document.getElementById('guarantor').value.trim();
        const guaranteeType = document.getElementById('guaranteeType').value;
        const reservedAmountGroup = document.getElementById('reservedAmountGroup');
        const reserveFundAmountGroup = document.getElementById('reserveFundAmountGroup');
        const remainingAmountGroup = document.getElementById('remainingAmountGroup');
        const reservedAmountRequired = document.getElementById('reservedAmountRequired');
        const reserveFundAmountRequired = document.getElementById('reserveFundAmountRequired');
        const remainingAmountRequired = document.getElementById('remainingAmountRequired');
        
        if (guaranteeType === '预留及备用金担保') {
            // 检查是否是备用金
            const isReserveFund = guarantor && (guarantor.includes('省区备用金') || guarantor.includes('大区备用金'));
            
            if (isReserveFund) {
                // 是备用金，显示备用金金额输入框和剩余可担保金额，隐藏预留金额
                if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'flex';
                if (reserveFundAmountRequired) reserveFundAmountRequired.style.display = 'inline';
                if (remainingAmountGroup) remainingAmountGroup.style.display = 'flex';
                if (remainingAmountRequired) remainingAmountRequired.style.display = 'inline';
                if (reservedAmountGroup) reservedAmountGroup.style.display = 'none';
            } else if (guarantor) {
                // 不是备用金，显示预留金额和剩余可担保金额，隐藏备用金金额
                if (reservedAmountGroup) reservedAmountGroup.style.display = 'flex';
                if (remainingAmountGroup) remainingAmountGroup.style.display = 'flex';
                if (reservedAmountRequired) reservedAmountRequired.style.display = 'inline';
                if (remainingAmountRequired) remainingAmountRequired.style.display = 'inline';
                if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'none';
            } else {
                // 没有选择担保方，都隐藏
                if (reservedAmountGroup) reservedAmountGroup.style.display = 'none';
                if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'none';
                if (remainingAmountGroup) remainingAmountGroup.style.display = 'none';
            }
        } else {
            // 不是预留及备用金担保，隐藏所有相关字段
            if (reservedAmountGroup) reservedAmountGroup.style.display = 'none';
            if (reserveFundAmountGroup) reserveFundAmountGroup.style.display = 'none';
            if (remainingAmountGroup) remainingAmountGroup.style.display = 'none';
        }
    }
    
    autoFillReservedAmount() {
        const guarantor = document.getElementById('guarantor').value.trim();
        const guaranteeType = document.getElementById('guaranteeType').value;
        
        if (guaranteeType === '预留及备用金担保' && guarantor) {
            // 检查是否是备用金
            const isReserveFund = guarantor.includes('省区备用金') || guarantor.includes('大区备用金');
            
            if (!isReserveFund) {
                // 不是备用金，从预留数据中查找对应担保方的预留金额
                // 支持多人，累加所有匹配的预留金额
                const guarantorNames = guarantor.split(/[,，]/).map(name => name.trim()).filter(name => name);
                const details = [];
                let totalReservedAmount = 0;
                
                for (const name of guarantorNames) {
                    const reserveRecord = this.reserveData.find(r => 
                        r.guarantor && r.guarantor.trim() === name
                    );
                    
                    if (reserveRecord && reserveRecord.reservedAmount) {
                        const amount = parseFloat(reserveRecord.reservedAmount) || 0;
                        totalReservedAmount += amount;
                        details.push({
                            name: name,
                            amount: amount
                        });
                    } else {
                        // 如果没有找到预留数据，也显示在明细中
                        details.push({
                            name: name,
                            amount: 0
                        });
                    }
                }
                
                // 更新预留金额
                const reservedAmountInput = document.getElementById('reservedAmount');
                if (reservedAmountInput) {
                    reservedAmountInput.value = totalReservedAmount;
                }
                
                // 显示明细
                const detailsContainer = document.getElementById('guarantorReserveDetails');
                const detailsList = document.getElementById('guarantorReserveList');
                if (detailsContainer && detailsList) {
                    if (details.length > 0) {
                        detailsList.innerHTML = details.map(d => {
                            const amountText = d.amount > 0 ? `¥${d.amount.toFixed(2)}` : '（无预留数据）';
                            return `<div style="margin: 2px 0;">${this.escapeHtml(d.name)}：${amountText}</div>`;
                        }).join('');
                        detailsContainer.style.display = 'block';
                    } else {
                        detailsContainer.style.display = 'none';
                    }
                }
                
                // 自动计算剩余可担保金额
                this.calculateRemainingAmount();
            } else {
                // 是备用金，隐藏明细
                const detailsContainer = document.getElementById('guarantorReserveDetails');
                if (detailsContainer) {
                    detailsContainer.style.display = 'none';
                }
            }
        } else {
            // 隐藏明细
            const detailsContainer = document.getElementById('guarantorReserveDetails');
            if (detailsContainer) {
                detailsContainer.style.display = 'none';
            }
        }
    }
    
    calculateRemainingAmount() {
        const guaranteeType = document.getElementById('guaranteeType')?.value;
        
        if (guaranteeType === '预留及备用金担保') {
            const guarantor = document.getElementById('guarantor')?.value.trim() || '';
            const isReserveFund = guarantor.includes('省区备用金') || guarantor.includes('大区备用金');
            
            const guaranteeAmountInput = document.getElementById('guaranteeAmount');
            const reservedPaymentAmountInput = document.getElementById('reservedPaymentAmount');
            const remainingAmountInput = document.getElementById('remainingAmount');
            
            if (isReserveFund) {
                // 备用金：使用备用金金额计算
                const reserveFundAmountInput = document.getElementById('reserveFundAmount');
                if (reserveFundAmountInput && guaranteeAmountInput && remainingAmountInput) {
                    const reserveFundAmount = parseFloat(reserveFundAmountInput.value) || 0;
                    const guaranteeAmount = parseFloat(guaranteeAmountInput.value) || 0;
                    const reservedPaymentAmount = parseFloat(reservedPaymentAmountInput?.value) || 0;
                    
                    // 剩余可担保金额 = 备用金金额 - 担保金额 + 预留回款金额
                    const remainingAmount = reserveFundAmount - guaranteeAmount + reservedPaymentAmount;
                    remainingAmountInput.value = Math.max(0, remainingAmount).toFixed(2);
                }
            } else {
                // 预留：使用预留金额计算
                const reservedAmountInput = document.getElementById('reservedAmount');
                if (reservedAmountInput && guaranteeAmountInput && remainingAmountInput) {
                    const reservedAmount = parseFloat(reservedAmountInput.value) || 0;
                    const guaranteeAmount = parseFloat(guaranteeAmountInput.value) || 0;
                    const reservedPaymentAmount = parseFloat(reservedPaymentAmountInput?.value) || 0;
                    
                    // 剩余可担保金额 = 预留金额 - 担保金额 + 预留回款金额
                    const remainingAmount = reservedAmount - guaranteeAmount + reservedPaymentAmount;
                    remainingAmountInput.value = Math.max(0, remainingAmount).toFixed(2);
                }
            }
        }
    }
    
    async importReserveData() {
        if (!window.XLSX) {
            this.showError('XLSX 库未加载，无法导入 Excel 文件');
            return;
        }
        
        // 创建隐藏的文件输入元素
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.style.display = 'none';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) {
                document.body.removeChild(input);
                return;
            }
            
            // 使用 FileReader 读取文件
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    // 解析数据（假设第一行是表头）
                    if (jsonData.length < 2) {
                        this.showError('Excel 文件数据不足，至少需要表头和一行数据');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 查找表头行
                    const headers = jsonData[0];
                    
                    // 查找关键列索引：拓展员列和最终合计列
                    const guarantorIndex = headers.findIndex(h => 
                        h && (h.toString().includes('拓展员') || h.toString().includes('担保方') || h.toString().includes('担保人'))
                    );
                    const reservedAmountIndex = headers.findIndex(h => 
                        h && (h.toString().includes('最终合计') || h.toString().includes('预留金额') || h.toString().includes('预留'))
                    );
                    
                    if (guarantorIndex === -1 || reservedAmountIndex === -1) {
                        this.showError('Excel 文件格式不正确，请确保包含"拓展员"列和"最终合计"列');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 解析数据行
                    const reserveRecords = [];
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0) continue;
                        
                        const guarantor = row[guarantorIndex] ? String(row[guarantorIndex]).trim() : '';
                        const reservedAmount = parseFloat(row[reservedAmountIndex]) || 0;
                        
                        // 支持拓展员名称、大区备用金、省区备用金
                        if (guarantor && (reservedAmount > 0 || guarantor.includes('备用金'))) {
                            reserveRecords.push({
                                guarantor: guarantor,
                                reservedAmount: reservedAmount > 0 ? reservedAmount : 0
                            });
                        }
                    }
                    
                    if (reserveRecords.length === 0) {
                        this.showError('未能从 Excel 文件中解析出有效的预留数据');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 保存预留数据
                    this.reserveData = reserveRecords;
                    this.saveReserveData();
                    
                    this.showToast(`✅ 成功导入 ${reserveRecords.length} 条预留数据`);
                    
                    // 如果当前表单是预留及备用金担保类型，自动填充预留金额
                    const currentGuaranteeType = document.getElementById('guaranteeType')?.value;
                    if (currentGuaranteeType === '预留及备用金担保') {
                        this.autoFillReservedAmount();
                    }
                    
                    // 清理
                    document.body.removeChild(input);
                } catch (err) {
                    console.error('解析 Excel 文件失败:', err);
                    this.showError('解析 Excel 文件失败: ' + err.message);
                    document.body.removeChild(input);
                }
            };
            
            reader.onerror = () => {
                this.showError('读取文件失败');
                document.body.removeChild(input);
            };
            
            reader.readAsArrayBuffer(file);
        };
        
        // 触发文件选择
        document.body.appendChild(input);
        input.click();
    }
    
    async importHistoryData() {
        if (!window.XLSX) {
            this.showError('XLSX 库未加载，无法导入 Excel 文件');
            return;
        }
        
        // 创建隐藏的文件输入元素
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.style.display = 'none';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) {
                document.body.removeChild(input);
                return;
            }
            
            // 使用 FileReader 读取文件
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    if (jsonData.length < 2) {
                        this.showError('Excel 文件数据不足，至少需要表头和一行数据');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 查找表头行
                    const headers = jsonData[0].map(h => h ? String(h).trim() : '');
                    
                    // 字段映射表：Excel表头 -> 数据字段名
                    const fieldMapping = {
                        '登记时间': 'registerTime',
                        '担保方': 'guarantor',
                        '担保类型': 'guaranteeType',
                        '对接人': 'contactPerson',
                        '大区': 'region',
                        '省区': 'province',
                        '离职标识': 'resignationFlag',
                        '是否使用工资担保': 'useSalaryGuarantee',
                        '是否使用提成担保': 'useCommissionGuarantee',
                        '工资提成担保金额': 'salaryCommissionAmount',
                        '预留金额': 'reservedAmount',
                        '备用金金额': 'reserveFundAmount',
                        '担保金额': 'guaranteeAmount',
                        '剩余可担保金额': 'remainingAmount',
                        '担保客户': 'guaranteedCustomer',
                        '订单号': 'orderNumber',
                        '预计回款时间': 'expectedPaymentTime',
                        '审批编号': 'approvalNumber',
                        '备注': 'remarks',
                        '工资提成回款': 'salaryCommissionPayment',
                        '预留回款金额': 'reservedPaymentAmount',
                        '未回款金额': 'unpaidAmount',
                        '回款时间': 'paymentTime'
                    };
                    
                    // 创建字段索引映射
                    const fieldIndexMap = {};
                    headers.forEach((header, index) => {
                        if (fieldMapping[header]) {
                            fieldIndexMap[fieldMapping[header]] = index;
                        }
                    });
                    
                    // 检查必要字段
                    if (!fieldIndexMap['registerTime'] && !fieldIndexMap['guarantor']) {
                        this.showError('Excel 文件格式不正确，请确保包含"登记时间"或"担保方"列');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 解析数据行
                    const importedRecords = [];
                    let successCount = 0;
                    let errorCount = 0;
                    const errors = [];
                    
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0) continue;
                        
                        try {
                            const record = {
                                id: Date.now() + i + Math.random(), // 生成唯一ID
                            };
                            
                            // 映射字段
                            Object.keys(fieldIndexMap).forEach(field => {
                                const colIndex = fieldIndexMap[field];
                                if (colIndex !== undefined && row[colIndex] !== undefined && row[colIndex] !== null && row[colIndex] !== '') {
                                    let value = row[colIndex];
                                    
                                    // 根据字段类型进行转换
                                    if (['salaryCommissionAmount', 'reservedAmount', 'reserveFundAmount', 'guaranteeAmount', 
                                         'remainingAmount', 'salaryCommissionPayment', 'reservedPaymentAmount', 
                                         'unpaidAmount'].includes(field)) {
                                        record[field] = parseFloat(value) || 0;
                                    } else if (field === 'paymentTime') {
                                        // 回款时间字段：转换为年月格式（YYYY-MM）
                                        if (typeof value === 'number' && value > 1 && value < 1000000) {
                                            // Excel 日期序列号
                                            const excelEpoch = new Date(1899, 11, 30);
                                            const date = new Date(excelEpoch.getTime() + (value - 1) * 24 * 60 * 60 * 1000);
                                            if (!isNaN(date.getTime())) {
                                                const year = date.getFullYear();
                                                const month = String(date.getMonth() + 1).padStart(2, '0');
                                                record[field] = `${year}-${month}`;
                                            } else {
                                                record[field] = String(value).trim();
                                            }
                                        } else {
                                            // 字符串格式，转换为年月格式
                                            const strValue = String(value).trim();
                                            // 如果是 YYYY/MM/DD 或 YYYY-MM-DD 格式，提取年月
                                            if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(strValue)) {
                                                const parts = strValue.split(/[\/\-]/);
                                                const year = parts[0].padStart(4, '0');
                                                const month = parts[1].padStart(2, '0');
                                                record[field] = `${year}-${month}`;
                                            } else if (/^\d{4}[\/\-]\d{1,2}$/.test(strValue)) {
                                                // 如果已经是 YYYY/MM 或 YYYY-MM 格式
                                                const parts = strValue.split(/[\/\-]/);
                                                const year = parts[0].padStart(4, '0');
                                                const month = parts[1].padStart(2, '0');
                                                record[field] = `${year}-${month}`;
                                            } else {
                                                record[field] = strValue;
                                            }
                                        }
                                    } else if (['expectedPaymentTime', 'registerTime'].includes(field)) {
                                        // 其他日期字段：处理 Excel 日期序列号或字符串格式
                                        if (typeof value === 'number' && value > 1 && value < 1000000) {
                                            // Excel 日期序列号
                                            const excelEpoch = new Date(1899, 11, 30);
                                            const date = new Date(excelEpoch.getTime() + (value - 1) * 24 * 60 * 60 * 1000);
                                            if (!isNaN(date.getTime())) {
                                                record[field] = date.toISOString().split('T')[0];
                                            } else {
                                                record[field] = String(value).trim();
                                            }
                                        } else {
                                            // 字符串格式的日期，尝试转换
                                            const strValue = String(value).trim();
                                            // 如果是 YYYY/MM/DD 格式，转换为 YYYY-MM-DD
                                            if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(strValue)) {
                                                const parts = strValue.split(/[\/\-]/);
                                                const year = parts[0].padStart(4, '0');
                                                const month = parts[1].padStart(2, '0');
                                                const day = parts[2].padStart(2, '0');
                                                record[field] = `${year}-${month}-${day}`;
                                            } else {
                                                record[field] = strValue;
                                            }
                                        }
                                    } else if (field === 'useSalaryGuarantee' || field === 'useCommissionGuarantee') {
                                        // 布尔值或选项值
                                        record[field] = String(value).trim();
                                    } else {
                                        record[field] = String(value).trim();
                                    }
                                }
                            });
                            
                            // 验证必要字段
                            if (!record.registerTime && !record.guarantor) {
                                errorCount++;
                                errors.push(`第 ${i + 1} 行：缺少必要字段（登记时间或担保方）`);
                                continue;
                            }
                            
                            // 如果没有登记时间，使用当前日期
                            if (!record.registerTime) {
                                record.registerTime = this.getTodayDate();
                            }
                            
                            // 确保数值字段有默认值并转换为数字类型
                            ['salaryCommissionAmount', 'reservedAmount', 'reserveFundAmount', 'guaranteeAmount', 
                             'remainingAmount', 'salaryCommissionPayment', 'reservedPaymentAmount', 
                             'unpaidAmount'].forEach(field => {
                                if (record[field] === undefined || record[field] === null || record[field] === '') {
                                    record[field] = 0;
                                } else {
                                    // 确保是数字类型
                                    const numValue = parseFloat(record[field]);
                                    record[field] = isNaN(numValue) ? 0 : numValue;
                                }
                            });
                            
                            // 根据金额字段自动填充担保类型（数据加工：总是根据金额自动填充，覆盖Excel中的值）
                            const salaryCommissionAmount = parseFloat(record.salaryCommissionAmount) || 0;
                            const guaranteeAmount = parseFloat(record.guaranteeAmount) || 0;
                            
                            // 根据金额自动判断并填充担保类型（强制覆盖）
                            if (salaryCommissionAmount > 0 && guaranteeAmount > 0) {
                                // 两者都不为0，则为组合担保
                                record.guaranteeType = '组合担保';
                            } else if (salaryCommissionAmount > 0) {
                                // 工资提成担保金额不为0，担保类型填工资提成担保
                                record.guaranteeType = '工资提成担保';
                            } else if (guaranteeAmount > 0) {
                                // 担保金额不为0，担保类型填预留及备用金担保
                                record.guaranteeType = '预留及备用金担保';
                            } else {
                                // 如果两者都为0，设置为空字符串
                                record.guaranteeType = '';
                            }
                            
                            // 调试信息（可在控制台查看）
                            if (i <= 3) { // 只打印前3条记录
                                console.log(`第${i+1}行数据加工:`, {
                                    salaryCommissionAmount,
                                    guaranteeAmount,
                                    guaranteeType: record.guaranteeType
                                });
                            }
                            
                            importedRecords.push(record);
                            successCount++;
                        } catch (err) {
                            errorCount++;
                            errors.push(`第 ${i + 1} 行：${err.message}`);
                        }
                    }
                    
                    if (importedRecords.length === 0) {
                        this.showError('未能从 Excel 文件中解析出有效的担保数据');
                        document.body.removeChild(input);
                        return;
                    }
                    
                    // 最终验证：确保所有记录的担保类型都根据金额正确设置
                    importedRecords.forEach((record, idx) => {
                        const salaryCommissionAmount = parseFloat(record.salaryCommissionAmount) || 0;
                        const guaranteeAmount = parseFloat(record.guaranteeAmount) || 0;
                        
                        // 强制根据金额重新设置担保类型
                        if (salaryCommissionAmount > 0 && guaranteeAmount > 0) {
                            record.guaranteeType = '组合担保';
                        } else if (salaryCommissionAmount > 0) {
                            record.guaranteeType = '工资提成担保';
                        } else if (guaranteeAmount > 0) {
                            record.guaranteeType = '预留及备用金担保';
                        } else {
                            record.guaranteeType = record.guaranteeType || '';
                        }
                    });
                    
                    // 直接覆盖现有数据
                    this.data = importedRecords;
                    
                    // 验证数据：检查前几条记录的担保类型是否正确设置
                    console.log('导入后的数据验证（前3条）:');
                    this.data.slice(0, 3).forEach((record, idx) => {
                        console.log(`记录${idx + 1}:`, {
                            guarantor: record.guarantor,
                            guaranteeType: record.guaranteeType,
                            salaryCommissionAmount: record.salaryCommissionAmount,
                            guaranteeAmount: record.guaranteeAmount
                        });
                    });
                    
                    this.saveData();
                    
                    if (errorCount > 0 && errors.length > 0) {
                        console.warn('导入过程中的错误：', errors);
                        this.showToast(`✅ 已覆盖数据，成功导入 ${successCount} 条记录，${errorCount} 条失败。请查看控制台了解详情。`);
                    } else {
                        this.showToast(`✅ 已覆盖数据，成功导入 ${successCount} 条历史担保数据`);
                    }
                    
                    // 刷新表格
                    if (this.currentView === 'list') {
                    this.refreshTable();
                }
                    
                    // 清理
                    document.body.removeChild(input);
                } catch (err) {
                    console.error('解析 Excel 文件失败:', err);
                    this.showError('解析 Excel 文件失败: ' + err.message);
                    document.body.removeChild(input);
                }
            };
            
            reader.onerror = () => {
                this.showError('读取文件失败');
                document.body.removeChild(input);
            };
            
            reader.readAsArrayBuffer(file);
        };
        
        document.body.appendChild(input);
        input.click();
    }
    
    updateHistoryData(id) {
        const record = this.data.find(r => r.id === id);
        if (!record) return;
        
        // 创建修改历史数据的对话框
        const dialog = document.createElement('div');
        dialog.className = 'update-dialog';
        dialog.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 10000;
        `;
        
        const dialogContent = document.createElement('div');
        dialogContent.className = 'update-dialog-content';
        dialogContent.style.cssText = `
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            max-width: 600px;
            width: 90%;
            max-height: 80vh;
            overflow-y: auto;
        `;
        
        const isSalaryGuarantee = record.guaranteeType === '工资提成担保';
        const isReserveGuarantee = record.guaranteeType === '预留及备用金担保';
        
        // Build conditional HTML parts separately to avoid nested template literal issues
        const salaryGuaranteeHtml = isSalaryGuarantee ? 
            '<div class="form-group"><label>工资提成回款 <span class="required">*</span></label><input type="number" id="updateSalaryCommissionPayment" step="0.01" min="0" value="' + (record.salaryCommissionPayment || 0) + '"><small style="color: var(--text-muted); font-size: 0.8rem;">输入回款金额后，未回款金额将自动计算</small></div>' : '';
        const reserveGuaranteeHtml = isReserveGuarantee ? 
            '<div class="form-group"><label>预留回款金额 <span class="required">*</span></label><input type="number" id="updateReservedPaymentAmount" step="0.01" min="0" value="' + (record.reservedPaymentAmount || 0) + '"><small style="color: var(--text-muted); font-size: 0.8rem;">输入回款金额后，未回款金额和剩余可担保金额将自动计算</small></div>' : '';
        
        dialogContent.innerHTML = `
            <div class="update-dialog-header">
                <h3>修改历史数据</h3>
                <button class="btn-icon" id="closeUpdateDialog" style="background: transparent; border: none; color: var(--text-secondary); font-size: 1.5rem; cursor: pointer;">✕</button>
            </div>
            <div class="update-dialog-body">
                <div class="form-group">
                    <label>登记时间</label>
                    <input type="text" value="${this.escapeHtml(record.registerTime || '')}" readonly style="background: var(--bg-secondary);">
                </div>
                <div class="form-group">
                    <label>担保方</label>
                    <input type="text" value="${this.escapeHtml(record.guarantor || '')}" readonly style="background: var(--bg-secondary);">
                </div>
                <div class="form-group">
                    <label>担保类型</label>
                    <input type="text" value="${this.escapeHtml(record.guaranteeType || '')}" readonly style="background: var(--bg-secondary);">
                </div>
                <div class="form-group">
                    <label>离职标识</label>
                    <select id="updateResignationFlag">
                        <option value="">请选择</option>
                        <option value="是" ${record.resignationFlag === '是' ? 'selected' : ''}>是</option>
                        <option value="否" ${record.resignationFlag === '否' ? 'selected' : ''}>否</option>
                    </select>
                </div>
                ${salaryGuaranteeHtml}
                ${reserveGuaranteeHtml}
                <div class="form-group">
                    <label>回款时间</label>
                    <input type="month" id="updatePaymentTime" value="${this.formatPaymentTime(record.paymentTime || '')}" placeholder="选择年月，如：2025-10">
                </div>
            </div>
            <div class="update-dialog-footer" style="margin-top: 20px; display: flex; justify-content: flex-end; gap: 12px;">
                <button class="btn btn-secondary" id="cancelUpdateBtn">取消</button>
                <button class="btn btn-primary" id="saveUpdateBtn">保存</button>
            </div>
        `;
        
        dialog.appendChild(dialogContent);
        document.body.appendChild(dialog);
        
        // 绑定事件
        const closeBtn = dialog.querySelector('#closeUpdateDialog');
        const cancelBtn = dialog.querySelector('#cancelUpdateBtn');
        const saveBtn = dialog.querySelector('#saveUpdateBtn');
        const updatePaymentTimeInput = document.getElementById('updatePaymentTime');
        
        const closeDialog = () => {
            document.body.removeChild(dialog);
        };
        
        closeBtn.addEventListener('click', closeDialog);
        cancelBtn.addEventListener('click', closeDialog);
        
        saveBtn.addEventListener('click', () => {
            const resignationFlag = document.getElementById('updateResignationFlag').value;
            let paymentTime = '';
            if (updatePaymentTimeInput && updatePaymentTimeInput.value) {
                // 回款时间使用年月格式（YYYY-MM）
                paymentTime = this.formatPaymentTime(updatePaymentTimeInput.value);
            }
            
            // 更新记录
            const index = this.data.findIndex(r => r.id === id);
            if (index !== -1) {
                const updatedRecord = { ...this.data[index] };
                updatedRecord.resignationFlag = resignationFlag;
                updatedRecord.paymentTime = paymentTime;
                
                if (isSalaryGuarantee) {
                    const salaryCommissionPayment = parseFloat(document.getElementById('updateSalaryCommissionPayment').value) || 0;
                    updatedRecord.salaryCommissionPayment = salaryCommissionPayment;
                    
                    // 计算未回款金额
                    const salaryCommissionAmount = parseFloat(updatedRecord.salaryCommissionAmount) || 0;
                    updatedRecord.unpaidAmount = Math.max(0, salaryCommissionAmount - salaryCommissionPayment);
                }
                
                if (isReserveGuarantee) {
                    const reservedPaymentAmount = parseFloat(document.getElementById('updateReservedPaymentAmount').value) || 0;
                    updatedRecord.reservedPaymentAmount = reservedPaymentAmount;
                    
                    // 计算未回款金额
                    const guaranteeAmount = parseFloat(updatedRecord.guaranteeAmount) || 0;
                    updatedRecord.unpaidAmount = Math.max(0, guaranteeAmount - reservedPaymentAmount);
                    
                    // 计算剩余可担保金额
                    // 如果是备用金，使用备用金金额；否则使用预留金额
                    const guarantor = updatedRecord.guarantor || '';
                    const isReserveFund = guarantor.includes('省区备用金') || guarantor.includes('大区备用金');
                    
                    if (isReserveFund) {
                        // 备用金：剩余可担保金额 = 备用金金额 - 担保金额 + 预留回款金额
                        const reserveFundAmount = parseFloat(updatedRecord.reserveFundAmount) || 0;
                        const remainingAmount = reserveFundAmount - guaranteeAmount + reservedPaymentAmount;
                        updatedRecord.remainingAmount = remainingAmount;
                    } else {
                        // 预留：剩余可担保金额 = 预留金额 - 担保金额 + 预留回款金额
                        const reservedAmount = parseFloat(updatedRecord.reservedAmount) || 0;
                        const remainingAmount = reservedAmount - guaranteeAmount + reservedPaymentAmount;
                        updatedRecord.remainingAmount = remainingAmount;
                    }
                }
                
                this.data[index] = updatedRecord;
                this.saveData();
                this.showToast('✅ 历史数据已更新');
                if (this.currentView === 'list') {
                    this.refreshTable();
                }
                closeDialog();
            }
        });
    }
}

