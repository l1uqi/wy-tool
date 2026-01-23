

// 政策外开单分析页面
export class OutOfPolicyPage {
    constructor(app) {
        this.app = app;
        this.selectedFiles = [];
        this.currentResult = null;
    }

    // 将 Excel 日期数字转换为日期字符串 (格式: 2025/11/4)
    excelDateToString(excelDate) {
        if (typeof excelDate === 'string') {
            // 如果已经是字符串，直接返回
            return excelDate;
        }
        if (typeof excelDate !== 'number' || excelDate <= 0) {
            return '-';
        }
        // Excel 日期转换：基准日期 1899-12-30，减去 2 修正 Excel 的 1900 闰年错误
        const epoch = new Date(1899, 11, 30); // 1899-12-30
        epoch.setDate(epoch.getDate() + excelDate - 2);
        return `${epoch.getFullYear()}/${epoch.getMonth() + 1}/${epoch.getDate()}`;
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">📉</span>
                        政策外开单分析
                    </h1>
                    <p class="page-desc">
                        导入出库表，自动分析低于挂网价开单的记录
                    </p>
                </div>
                
                <div class="upload-section slide-up">
                    <div class="upload-group">
                        <div style="margin-bottom: 16px;">
                            <h3 style="margin-bottom: 8px;">1. 导入出库表</h3>
                            <p style="color: var(--text-secondary); font-size: 0.9rem; margin-bottom: 12px;">
                                请选择出库明细 Excel 文件
                            </p>
                            <button class="btn btn-primary" id="selectFilesBtn">
                                <span>📂</span> 选择文件
                            </button>
                        </div>
                        
                        <div id="fileList" class="file-list" style="display: none; margin-top: 16px; padding: 12px; background: var(--bg-secondary); border-radius: 8px;">
                            <!-- 文件列表将在这里显示 -->
                        </div>
                    </div>
                </div>
                
                <div class="result-section" id="resultSection" style="display: none; margin-top: 24px;">
                    <div class="result-header" style="margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center;">
                        <h2 class="result-title">数据列表 <span id="recordCount" style="font-size: 0.9rem; color: var(--text-muted); font-weight: normal;"></span></h2>
                    </div>
                    <div class="table-container" style="position: relative; background: var(--bg-card); border-radius: 8px; box-shadow: var(--shadow-sm);">
                        <!-- 固定表头 -->
                        <div id="tableHeader" style="overflow: hidden; background: var(--bg-secondary);">
                            <table style="width: 100%; border-collapse: collapse; white-space: nowrap;">
                                <thead>
                                    <tr style="text-align: left;">
                                        <th style="padding: 12px; min-width: 100px;">下单日期</th>
                                        <th style="padding: 12px; width: 200px; max-width: 200px;">客户名称</th>
                                        <th style="padding: 12px; min-width: 120px;">商品编码</th>
                                        <th style="padding: 12px; width: 150px; max-width: 150px;">通用名</th>
                                        <th style="padding: 12px; text-align: center; min-width: 80px;">低于挂网</th>
                                        <th style="padding: 12px; min-width: 150px;">活动政策</th>
                                    </tr>
                                </thead>
                            </table>
                        </div>
                        <!-- 可滚动表体 -->
                        <div id="tableBodyContainer" style="overflow: auto; max-height: 600px;">
                            <table style="width: 100%; border-collapse: collapse; white-space: nowrap;">
                                <tbody id="tableBody">
                                    <!-- 数据行 -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        this.bindEvents(container);
    }
    
    bindEvents(container) {
        const selectFilesBtn = container.querySelector('#selectFilesBtn');
        const analyzeBtn = container.querySelector('#analyzeBtn');
        
        if (selectFilesBtn) {
            selectFilesBtn.addEventListener('click', () => this.handleFileSelection());
        }
        
        if (analyzeBtn) {
            analyzeBtn.addEventListener('click', () => this.runAnalysis());
        }
    }
    
    async handleFileSelection() {
        if (!window.__TAURI__) {
            alert('Tauri API 不可用');
            return;
        }
        
        const { open } = window.__TAURI__.dialog;
        const { invoke } = window.__TAURI__.core;
        
        try {
            const selected = await open({
                multiple: false,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx', 'xls']
                }]
            });
            
            if (selected) {
                this.selectedFiles = [selected]; // Keep array format for compatibility
                this.updateFileList();
                
                // 立即加载文件
                this.loadAndDisplayData(selected);
            }
        } catch (error) {
            console.error('选择文件失败:', error);
            alert('选择文件失败: ' + error);
        }
    }

    async loadAndDisplayData(filePath) {
        const { invoke } = window.__TAURI__.core;
        const resultSection = document.getElementById('resultSection');
        const tableBody = document.getElementById('tableBody');
        const recordCount = document.getElementById('recordCount');

        if (!resultSection || !tableBody) return;

        try {
            // Show loading state
            resultSection.style.display = 'block';
            tableBody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 20px;">正在加载数据...</td></tr>';

            const result = await invoke('load_out_of_policy_excel', { filePath });

            if (result && result.rows) {
                this.currentResult = result;
                recordCount.textContent = `(共 ${result.total_rows} 条记录)`;

                // 虚拟列表配置
                this.virtualListConfig = {
                    itemHeight: 48, // 每行高度
                    bufferSize: 5, // 缓冲区行数
                    startIndex: 0,
                    endIndex: 0
                };

                // Render table rows with virtual scrolling
                this.renderVirtualTable(result.rows);

                // Keep result section visible after loading
                setTimeout(() => {
                    resultSection.style.display = 'block';
                }, 500);
            }
        } catch (error) {
            console.error('加载数据失败:', error);
            tableBody.innerHTML = `<tr><td colspan="10" style="text-align: center; padding: 20px; color: #ef4444;">加载失败: ${error}</td></tr>`;
        }
    }

    showExportButton() {
        const header = document.querySelector('.result-header');
        if (!header) return;

        // 检查是否已有导出按钮
        if (!document.getElementById('exportBtn')) {
            const exportBtn = document.createElement('button');
            exportBtn.id = 'exportBtn';
            exportBtn.className = 'btn btn-primary';
            exportBtn.innerHTML = '<span>📥</span> 导出Excel';
            exportBtn.style.marginLeft = '12px';
            exportBtn.addEventListener('click', () => this.exportData());
            header.appendChild(exportBtn);
        }
    }

    renderVirtualTable(rows) {
        const tableBodyContainer = document.getElementById('tableBodyContainer');
        const tableHeader = document.getElementById('tableHeader');
        const tableBody = document.getElementById('tableBody');

        if (!tableBodyContainer || !tableHeader || !tableBody) return;

        // 同步横向滚动
        tableBodyContainer.addEventListener('scroll', () => {
            tableHeader.scrollLeft = tableBodyContainer.scrollLeft;
        });

        // 计算需要渲染的行数
        const containerHeight = 600;
        const { bufferSize, itemHeight } = this.virtualListConfig;

        // 绑定滚动事件
        tableBodyContainer.onscroll = () => {
            this.updateVirtualRows(rows, tableBodyContainer);
        };

        // 初始渲染
        this.updateVirtualRows(rows, tableBodyContainer);
    }

    updateVirtualRows(rows, tableBodyContainer) {
        const { itemHeight, bufferSize } = this.virtualListConfig;
        const scrollTop = tableBodyContainer.scrollTop;

        // 计算可见范围的索引
        const startIndex = Math.max(0, Math.floor(scrollTop / itemHeight) - bufferSize);
        const endIndex = Math.min(
            rows.length,
            startIndex + Math.ceil(tableBodyContainer.clientHeight / itemHeight) + bufferSize * 2
        );

        this.virtualListConfig.startIndex = startIndex;
        this.virtualListConfig.endIndex = endIndex;

        // 渲染可见行
        const visibleRows = rows.slice(startIndex, endIndex).map((row, i) => {
            const isBelow = row.is_below_listed === '是' || row.is_below_listed === 'True' || row.is_below_listed === 'TRUE';

            return `
                <tr style="border-bottom: 1px solid var(--border-color);">
                    <td style="padding: 12px; min-width: 100px;">${this.excelDateToString(row.order_date)}</td>
                    <td style="padding: 12px; width: 200px; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.customer_name}">${row.customer_name}</td>
                    <td style="padding: 12px; min-width: 120px;">${row.product_code}</td>
                    <td style="padding: 12px; width: 150px; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.generic_name}">${row.generic_name}</td>
                    <td style="padding: 12px; text-align: center; min-width: 80px;">
                        <span style="padding: 2px 6px; border-radius: 4px; font-size: 0.85rem; background: ${isBelow ? 'rgba(239, 68, 68, 0.1)' : 'rgba(16, 185, 129, 0.1)'}; color: ${isBelow ? '#ef4444' : '#10b981'};">
                            ${row.is_below_listed}
                        </span>
                    </td>
                    <td style="padding: 12px; min-width: 150px;">${row.policy || '-'}</td>
                </tr>
            `;
        }).join('');

        // 添加占位 div 维持滚动高度
        const topSpacerHeight = startIndex * itemHeight;
        const bottomSpacerHeight = (rows.length - endIndex) * itemHeight;

        const tableBody = document.getElementById('tableBody');
        tableBody.innerHTML = `
            <tr><td colspan="6" style="padding: 0; height: ${topSpacerHeight}px; border: none;"></td></tr>
            ${visibleRows}
            <tr><td colspan="6" style="padding: 0; height: ${bottomSpacerHeight}px; border: none;"></td></tr>
        `;
    }

    async exportData() {
        if (!this.currentResult || !this.currentResult.rows) {
            alert('没有数据可导出');
            return;
        }

        if (!window.__TAURI__) {
            alert('Tauri API 不可用');
            return;
        }

        try {
            // 准备数据
            const headers = [
                '下单日期', '销售单号', '客户编码', '客户名称', '商品编码', '通用名',
                '销售单价/积分', '结算单价', '挂网价', '是否低于挂网', '是否活动政策内',
                '活动政策', '活动后底价', '毛利率(%)', '销售数量', '支付金额',
            ];

            const rows = this.currentResult.rows.map(row => [
                row.order_date,
                row.sales_order_no,
                row.customer_code,
                row.customer_name,
                row.product_code,
                row.generic_name,
                row.sales_price,
                row.settlement_price,
                row.listed_price,
                row.is_below_listed,
                row.is_in_policy,
                row.policy,
                row.base_price_after_policy,
                row.gross_margin_rate,
                row.sales_quantity,
                row.pay_amount,
            ]);

            // 创建工作簿
            const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '分析结果');

            // 生成Excel二进制数据
            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

            // 转换为base64
            const base64 = this.arrayBufferToBase64(excelBuffer);

            // 保存文件
            const { invoke } = window.__TAURI__.core;
            const { save } = window.__TAURI__.dialog;

            const filePath = await save({
                defaultPath: `分析结果_已匹配_${new Date().toISOString().slice(0,10).replace(/-/g, '')}.xlsx`,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx']
                }]
            });

            if (filePath) {
                await invoke('save_excel_file', {
                    filePath: filePath,
                    content: base64
                });
                alert('导出成功!');
            }
        } catch (error) {
            console.error('导出失败:', error);
            alert(`导出失败: ${error}`);
        }
    }

    arrayBufferToBase64(buffer) {
        let binary = '';
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return window.btoa(binary);
    }
    
    updateFileList() {
        const fileList = document.getElementById('fileList');
        
        if (!fileList) return;
        
        if (this.selectedFiles.length === 0) {
            fileList.style.display = 'none';
            return;
        }
        
        fileList.style.display = 'block';
        fileList.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <span style="font-weight: 500;">已选择文件:</span>
                <button class="btn btn-sm btn-danger" id="clearFilesBtn" style="padding: 2px 8px; font-size: 0.8rem;">重新选择</button>
            </div>
            <ul style="list-style: none; padding: 0; margin: 0;">
                ${this.selectedFiles.map(file => `
                    <li style="padding: 6px 0; border-bottom: 1px solid var(--border-color); color: var(--text-secondary); font-size: 0.9rem; word-break: break-all;">
                        📄 ${file.split(/[/\\]/).pop()}
                    </li>
                `).join('')}
            </ul>
        `;
        
        const clearBtn = document.getElementById('clearFilesBtn');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => {
                this.selectedFiles = [];
                this.updateFileList();
                document.getElementById('resultSection').style.display = 'none';
            });
        }
    }

}