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
                            <button class="btn btn-primary" id="importDataSourceBtn">
                                <span>📥</span>
                                添加数据源
                            </button>
                        </div>
                        <div class="data-source-list" id="dataSourceList">
                            <p class="data-source-hint" id="emptyHint">
                                暂无数据源，请点击"添加数据源"按钮导入 Excel 文件
                            </p>
                            <div class="data-source-items" id="dataSourceItems"></div>
                            <div class="import-progress" id="importProgress" style="display: none;">
                                <div class="progress-bar-container">
                                    <div class="progress-bar" id="progressBar"></div>
                                </div>
                                <p class="progress-text" id="progressText">正在导入数据源...</p>
                            </div>
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
                    
                    <div class="feature-card slide-up" data-feature="rebate">
                        <div class="feature-icon green">✨</div>
                        <h3 class="feature-title">返利计算</h3>
                        <p class="feature-desc">
                            支持区间挂网底价的返利计算，导入挂网底价表和订单表即可快速计算返利，
                            自动处理复杂的返利规则和条件判断
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
        const progressContainer = document.getElementById('importProgress');
        const progressBar = document.getElementById('progressBar');
        const progressText = document.getElementById('progressText');
        
        if (progressContainer && progressBar && progressText) {
            progressContainer.style.display = 'block';
            const percent = Math.round(progress.percent || 0);
            progressBar.style.width = `${percent}%`;
            progressText.textContent = `正在导入数据源... ${percent}%`;
        }
    }
    
    hideProgress() {
        const progressContainer = document.getElementById('importProgress');
        if (progressContainer) {
            progressContainer.style.display = 'none';
        }
    }
    
    async loadDataSourceInfo() {
        if (!window.__TAURI__) {
            console.warn('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            console.log('开始加载数据源列表...');
            const listInfo = await invoke('get_data_source_list_info');
            console.log('获取到的数据源列表:', listInfo);
            this.updateDataSourceUI(listInfo);
        } catch (error) {
            console.error('加载数据源信息失败:', error);
            this.updateDataSourceUI({ data_sources: [], current_id: null });
        }
    }
    
    updateDataSourceUI(listInfo) {
        console.log('更新UI，数据源列表:', listInfo);
        
        const itemsContainer = document.getElementById('dataSourceItems');
        const emptyHint = document.getElementById('emptyHint');
        
        if (!itemsContainer || !emptyHint) {
            console.error('DOM元素未找到:', { 
                itemsContainer: !!itemsContainer, 
                emptyHint: !!emptyHint
            });
            // 如果元素不存在，可能是页面还没渲染完成，延迟重试
            setTimeout(() => {
                console.log('延迟重试更新UI...');
                this.updateDataSourceUI(listInfo);
            }, 200);
            return;
        }
        
        if (!listInfo || !listInfo.data_sources || listInfo.data_sources.length === 0) {
            console.log('没有数据源，显示空提示');
            emptyHint.style.display = 'block';
            itemsContainer.innerHTML = '';
            return;
        }
        
        console.log(`显示 ${listInfo.data_sources.length} 个数据源`);
        emptyHint.style.display = 'none';
        
        this.renderDataSourceItems(itemsContainer, listInfo);
    }
    
    renderDataSourceItems(itemsContainer, listInfo) {
        
        itemsContainer.innerHTML = listInfo.data_sources.map(ds => {
            const isCurrent = listInfo.current_id === ds.id;
            return `
                <div class="data-source-item ${isCurrent ? 'current' : ''}" data-id="${ds.id}">
                    <div class="ds-item-main">
                        <div class="ds-item-info">
                            <div class="ds-item-name">
                                ${isCurrent ? '<span class="current-badge">当前</span>' : ''}
                                <strong>${this.escapeHtml(ds.file_name)}</strong>
                            </div>
                            <div class="ds-item-meta">
                                <span>${ds.total_rows.toLocaleString()} 行</span>
                                <span>•</span>
                                <span>${ds.loaded_at}</span>
                            </div>
                        </div>
                        <div class="ds-item-actions">
                            ${!isCurrent ? `
                                <button class="btn btn-sm btn-primary switch-ds-btn" data-id="${ds.id}">
                                    <span>✓</span> 使用
                                </button>
                            ` : ''}
                            <button class="btn btn-sm btn-danger delete-ds-btn" data-id="${ds.id}">
                                <span>🗑️</span> 删除
                            </button>
                        </div>
                    </div>
                </div>
            `;
        }).join('');
        
        console.log('数据源列表HTML已更新，绑定事件...');
        
        // 绑定事件
        itemsContainer.querySelectorAll('.switch-ds-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const button = e.currentTarget;
                const id = button.dataset.id;
                if (id) {
                    this.switchDataSource(id);
                }
            });
        });
        
        itemsContainer.querySelectorAll('.delete-ds-btn').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                e.stopPropagation();
                e.preventDefault();
                const button = e.currentTarget;
                const id = button.dataset.id;
                if (id) {
                    await this.deleteDataSource(id);
                }
            });
        });
        
        console.log('事件绑定完成');
    }
    
    async switchDataSource(id) {
        if (!window.__TAURI__) return;
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            await invoke('switch_data_source', { dataSourceId: id });
            this.showToast('✅ 已切换到该数据源');
            await this.loadDataSourceInfo();
        } catch (error) {
            this.showError('切换数据源失败: ' + error);
        }
    }
    
    async deleteDataSource(id) {
        if (!window.__TAURI__) {
            this.showError('Tauri API 不可用');
            return;
        }
        
        const { invoke } = window.__TAURI__.core;
        
        try {
            await invoke('delete_data_source', { dataSourceId: id });
            this.showToast('✅ 数据源已删除');
            await this.loadDataSourceInfo();
        } catch (error) {
            this.showError('删除数据源失败: ' + error);
        }
    }
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
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
                } else if (feature === 'rebate') {
                    window.location.hash = 'rebate';
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
                console.log('选择的文件:', selected);
                
                // 显示导入进度提示
                this.showImportProgress();
                
                try {
                    console.log('开始添加数据源...');
                    const result = await invoke('add_data_source', { filePath: selected });
                    console.log('添加数据源结果:', result);
                    
                    // 隐藏进度提示
                    this.hideProgress();
                    
                    this.showToast(`✅ 数据源添加成功！共 ${result.total_rows.toLocaleString()} 行数据`);
                    
                    // 等待一下确保后端保存完成
                    await new Promise(resolve => setTimeout(resolve, 200));
                    
                    // 重新加载数据源列表
                    console.log('重新加载数据源列表...');
                    try {
                        await this.loadDataSourceInfo();
                        console.log('数据源列表已刷新');
                        
                        // 再次检查，确保UI已更新
                        setTimeout(() => {
                            const itemsContainer = document.getElementById('dataSourceItems');
                            const emptyHint = document.getElementById('emptyHint');
                            console.log('UI检查:', {
                                itemsContainer: !!itemsContainer,
                                emptyHint: !!emptyHint,
                                itemsCount: itemsContainer?.children.length || 0,
                                emptyHintDisplay: emptyHint?.style.display
                            });
                        }, 300);
                    } catch (error) {
                        console.error('刷新数据源列表失败:', error);
                        // 即使失败也尝试重新加载
                        setTimeout(() => this.loadDataSourceInfo(), 500);
                    }
                } catch (error) {
                    console.error('添加数据源失败:', error);
                    // 隐藏进度提示
                    this.hideProgress();
                    this.showError('添加数据源失败: ' + error);
                }
            } else {
                console.log('用户取消了文件选择');
            }
        } catch (error) {
            console.error('选择文件失败:', error);
            this.showError('选择文件失败: ' + error);
        }
    }
    
    showImportProgress() {
        const progressContainer = document.getElementById('importProgress');
        const progressBar = document.getElementById('progressBar');
        const progressText = document.getElementById('progressText');
        
        if (progressContainer && progressBar && progressText) {
            progressContainer.style.display = 'block';
            progressBar.style.width = '0%';
            progressText.textContent = '正在导入数据源，请稍候...';
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

