

// 政策外开单分析页面
export class OutOfPolicyPage {
    constructor(app) {
        this.app = app;
        this.selectedFiles = [];
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
                        导入出库表，自动匹配活动政策并导出
                    </p>
                </div>

                <div class="upload-section slide-up">
                    <div class="upload-group">
                        <div style="margin-bottom: 16px;">
                            <h3 style="margin-bottom: 8px;">导入Excel文件</h3>
                            <p style="color: var(--text-secondary); font-size: 0.9rem; margin-bottom: 12px;">
                                请选择包含"分析"和"2025活动政策"两个工作表的 Excel 文件
                            </p>
                            <button class="btn btn-primary" id="selectFilesBtn">
                                <span>📂</span> 选择文件并处理
                            </button>
                        </div>

                        <div id="fileList" class="file-list" style="display: none; margin-top: 16px; padding: 12px; background: var(--bg-secondary); border-radius: 8px;">
                            <!-- 状态显示 -->
                        </div>
                    </div>
                </div>
            </div>
        `;

        this.bindEvents(container);
    }

    bindEvents(container) {
        const selectFilesBtn = container.querySelector('#selectFilesBtn');

        if (selectFilesBtn) {
            selectFilesBtn.addEventListener('click', () => this.handleFileSelection());
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
                this.processFile(selected);
            }
        } catch (error) {
            console.error('选择文件失败:', error);
            alert('选择文件失败: ' + error);
        }
    }

    async processFile(filePath) {
        const { invoke } = window.__TAURI__.core;
        const fileList = document.getElementById('fileList');

        if (!fileList) return;

        try {
            // 显示加载状态
            fileList.style.display = 'block';
            fileList.innerHTML = '<div style="text-align: center; padding: 20px;">正在处理数据...</div>';

            const result = await invoke('load_out_of_policy_excel', { filePath });

            if (result) {
                // 文件已直接生成，显示成功信息
                fileList.innerHTML = `
                    <div style="text-align: center; padding: 20px;">
                        <div style="color: #10b981; margin-bottom: 12px;">处理成功!</div>
                        <div style="color: var(--text-secondary); font-size: 0.9rem;">
                            已生成文件: ${result.file_path.split(/[/\\]/).pop()}
                        </div>
                        <div style="color: var(--text-secondary); font-size: 0.85rem; margin-top: 8px;">
                            共处理 ${result.total_rows} 条记录
                        </div>
                    </div>
                `;
            }
        } catch (error) {
            console.error('处理失败:', error);
            fileList.innerHTML = `<div style="text-align: center; padding: 20px; color: #ef4444;">处理失败: ${error}</div>`;
        }
    }
}
