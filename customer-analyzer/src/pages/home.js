// 首页组件
export class HomePage {
    constructor(app) {
        this.app = app;
        this.dataSourceInfo = null;
        this.unlistenProgress = null;
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="home-container">
                <div class="home-hero slide-up">
                    <h1 class="home-title">婉怡的工具箱</h1>
                    <p class="home-subtitle">
                        专为财务工作打造的数据分析工具，快速处理大型 Excel 财务数据，
                        让繁琐的数据统计工作变得轻松高效
                    </p>
                </div>
                
                <!-- 数据源管理区域 -->
                <div class="data-source-section slide-up">
                    <div class="data-source-card">
                        <div class="data-source-header">
                            <h3>📁 数据源管理</h3>
                            <span class="data-source-status" id="dataSourceStatus">未导入</span>
                        </div>
                        <div class="data-source-content" id="dataSourceContent">
                            <p class="data-source-hint">请先导入 Excel 数据源，所有分析功能将使用此数据源</p>
                            <button class="btn btn-primary" id="importDataSourceBtn">
                                <span>📥</span>
                                导入数据源
                            </button>
                        </div>
                        <div class="data-source-info" id="dataSourceInfo" style="display: none;">
                            <div class="info-row">
                                <span class="info-label">文件名：</span>
                                <span class="info-value" id="dataSourceFileName">-</span>
                            </div>
                            <div class="info-row">
                                <span class="info-label">数据行数：</span>
                                <span class="info-value" id="dataSourceRows">-</span>
                            </div>
                            <div class="info-row">
                                <span class="info-label">导入时间：</span>
                                <span class="info-value" id="dataSourceTime">-</span>
                            </div>
                            <button class="btn btn-secondary" id="changeDataSourceBtn">
                                <span>🔄</span>
                                更换数据源
                            </button>
                        </div>
                    </div>
                </div>
                
                <div class="features-grid">
                    <div class="feature-card slide-up" data-feature="top20">
                        <div class="feature-icon blue">📊</div>
                        <h3 class="feature-title">前20大客户分析</h3>
                        <p class="feature-desc">
                            快速分析财务 Excel 数据，自动汇总客户应收金额，
                            生成前20大客户账务排行榜，轻松处理大量财务数据
                        </p>
                        <span class="feature-badge rust">⚡ Rust高性能</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="monthly">
                        <div class="feature-icon purple">📈</div>
                        <h3 class="feature-title">月度财务分析</h3>
                        <p class="feature-desc">
                            按客户或地区统计每月财务数据，生成可视化趋势图表，
                            支持环比增长分析，让财务数据一目了然
                        </p>
                        <span class="feature-badge new">✨ 新功能</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon cyan">🗂️</div>
                        <h3 class="feature-title">客户账务分类</h3>
                        <p class="feature-desc">
                            基于应收金额和账期自动将客户分为 VIP、普通、
                            潜力等类别，便于财务管理和风险控制
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon green">📋</div>
                        <h3 class="feature-title">财务报表生成</h3>
                        <p class="feature-desc">
                            一键生成专业的财务分析报表，支持多种格式导出，
                            让财务汇报工作更轻松
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon amber">🔄</div>
                        <h3 class="feature-title">财务数据对比</h3>
                        <p class="feature-desc">
                            多文件财务数据对比分析，快速发现账务差异和变化，
                            便于财务核对和审计
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon rose">⚙️</div>
                        <h3 class="feature-title">批量账务处理</h3>
                        <p class="feature-desc">
                            批量导入多个财务 Excel 文件，自动合并处理，
                            大幅提升财务工作效率
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                </div>
            </div>
        `;
        
        await this.loadDataSourceInfo();
        this.bindEvents(container);
        this.setupProgressListener();
    }
    
    async setupProgressListener() {
        if (!window.__TAURI__) return;
        
        const { listen } = window.__TAURI__.event;
        
        if (this.unlistenProgress) {
            this.unlistenProgress();
        }
        
        this.unlistenProgress = await listen('excel-progress', (event) => {
            const progress = event.payload;
            this.updateProgress(progress);
        });
    }
    
    updateProgress(progress) {
        const statusEl = document.getElementById('dataSourceStatus');
        if (statusEl) {
            statusEl.textContent = `导入中 ${progress.percent}%`;
        }
    }
    
    async loadDataSourceInfo() {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            const info = await invoke('get_data_source_info');
            if (info) {
                this.dataSourceInfo = info;
                this.updateDataSourceUI(info);
            } else {
                this.updateDataSourceUI(null);
            }
        } catch (error) {
            console.error('加载数据源信息失败:', error);
            this.updateDataSourceUI(null);
        }
    }
    
    updateDataSourceUI(info) {
        const contentEl = document.getElementById('dataSourceContent');
        const infoEl = document.getElementById('dataSourceInfo');
        const statusEl = document.getElementById('dataSourceStatus');
        
        if (info && info.file_path) {
            // 显示数据源信息
            contentEl.style.display = 'none';
            infoEl.style.display = 'block';
            document.getElementById('dataSourceFileName').textContent = info.file_name || '未知文件';
            document.getElementById('dataSourceRows').textContent = 
                info.total_rows > 0 ? info.total_rows.toLocaleString() + ' 行' : '未加载';
            document.getElementById('dataSourceTime').textContent = info.loaded_at || '-';
            statusEl.textContent = '已导入';
            statusEl.className = 'data-source-status status-loaded';
        } else {
            // 显示导入提示
            contentEl.style.display = 'block';
            infoEl.style.display = 'none';
            statusEl.textContent = '未导入';
            statusEl.className = 'data-source-status status-empty';
        }
    }
    
    bindEvents(container) {
        const importBtn = container.querySelector('#importDataSourceBtn');
        const changeBtn = container.querySelector('#changeDataSourceBtn');
        
        if (importBtn) {
            importBtn.addEventListener('click', () => this.importDataSource());
        }
        
        if (changeBtn) {
            changeBtn.addEventListener('click', () => this.importDataSource());
        }
        
        container.querySelectorAll('.feature-card').forEach(card => {
            card.addEventListener('click', () => {
                const feature = card.dataset.feature;
                if (feature === 'top20') {
                    window.location.hash = 'top20';
                } else if (feature === 'monthly') {
                    window.location.hash = 'monthly';
                } else if (feature === 'coming') {
                    this.showToast('该功能正在开发中，敬请期待！');
                }
            });
        });
    }
    
    async importDataSource() {
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        const { open } = window.__TAURI__.dialog;
        
        try {
            const selected = await open({
                multiple: false,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx', 'xls']
                }]
            });
            
            if (selected) {
                const statusEl = document.getElementById('dataSourceStatus');
                statusEl.textContent = '导入中...';
                statusEl.className = 'data-source-status status-loading';
                
                try {
                    const result = await invoke('set_data_source', { filePath: selected });
                    
                    this.dataSourceInfo = {
                        file_path: result.file_path,
                        file_name: result.file_name,
                        loaded_at: new Date().toLocaleString('zh-CN'),
                        total_rows: result.total_rows,
                    };
                    
                    this.updateDataSourceUI(this.dataSourceInfo);
                    this.showToast(`✅ 数据源导入成功！共 ${result.total_rows.toLocaleString()} 行数据`);
                } catch (error) {
                    statusEl.textContent = '导入失败';
                    statusEl.className = 'data-source-status status-error';
                    this.showError('导入失败: ' + error);
                }
            }
        } catch (error) {
            console.error('选择文件失败:', error);
            this.showError('选择文件失败: ' + error);
        }
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
            border: 1px solid var(--border-color);
            color: var(--text-primary);
            padding: 16px 32px;
            border-radius: 12px;
            box-shadow: var(--shadow-lg);
            z-index: 1000;
            animation: toastIn 0.3s ease;
        `;
        toast.textContent = message;
        
        if (!document.querySelector('#toast-styles')) {
            const style = document.createElement('style');
            style.id = 'toast-styles';
            style.textContent = `
                @keyframes toastIn {
                    from {
                        opacity: 0;
                        transform: translateX(-50%) translateY(20px);
                    }
                    to {
                        opacity: 1;
                        transform: translateX(-50%) translateY(0);
                    }
                }
                @keyframes toastOut {
                    from {
                        opacity: 1;
                        transform: translateX(-50%) translateY(0);
                    }
                    to {
                        opacity: 0;
                        transform: translateX(-50%) translateY(20px);
                    }
                }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'toastOut 0.3s ease forwards';
            setTimeout(() => toast.remove(), 300);
        }, 2500);
    }
}
