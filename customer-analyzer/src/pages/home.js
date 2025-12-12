// 首页组件
export class HomePage {
    constructor(app) {
        this.app = app;
        this.dataSourceInfo = null;
        this.unlistenProgress = null;
        this.currentImportingFileIndex = null;
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
                                <p class="progress-detail" id="progressDetail" style="margin-top: 8px; font-size: 0.9rem; color: var(--text-muted);"></p>
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
                    
                    <div class="feature-card slide-up" data-feature="purchase">
                        <div class="feature-icon purple">💰</div>
                        <h3 class="feature-title">客户采购额计算</h3>
                        <p class="feature-desc">
                            导入客户编码表，自动匹配数据源中的客户数据，按月份统计每个客户的采购金额，
                            支持多数据源合并分析，生成详细的月度采购报表
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
        // 进度监听器已在 App 级别设置，这里不需要重复设置
        // 但保留方法以兼容现有代码
    }
    
    updateProgress(progress) {
        const percent = Math.round(progress.percent || 0);
        
        // 如果有当前导入的文件索引（串行导入），更新该文件的进度条
        if (this.currentImportingFileIndex !== null && this.currentImportingFileIndex !== undefined) {
            if (window.app && window.app.updateFileProgress) {
                window.app.updateFileProgress(
                    this.currentImportingFileIndex,
                    percent,
                    '导入中',
                    `${percent}%`
                );
            }
        } else {
            // 单个文件导入，使用全局进度显示
            if (window.app && window.app.updateGlobalProgress) {
                window.app.updateGlobalProgress(progress);
            }
        }
    }
    
    updateProgressOverall(percent, text) {
        // 使用全局进度显示
        if (window.app && window.app.updateGlobalProgressOverall) {
            window.app.updateGlobalProgressOverall(percent, text);
        }
    }
    
    updateProgressDetail(detail) {
        // 使用全局进度显示
        if (window.app && window.app.updateGlobalProgressDetail) {
            window.app.updateGlobalProgressDetail(detail);
        }
    }
    
    hideProgress() {
        // 使用全局进度显示
        if (window.app && window.app.hideGlobalProgress) {
            window.app.hideGlobalProgress();
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
            
            // 确保UI更新
            this.updateDataSourceUI(listInfo);
            
            // 如果UI更新失败，延迟重试一次
            setTimeout(() => {
                const itemsContainer = document.getElementById('dataSourceItems');
                if (!itemsContainer || itemsContainer.children.length === 0) {
                    if (listInfo && listInfo.data_sources && listInfo.data_sources.length > 0) {
                        console.log('UI更新可能失败，重试更新...');
                        this.updateDataSourceUI(listInfo);
                    }
                }
            }, 100);
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
            return `
                <div class="data-source-item" data-id="${ds.id}">
                    <div class="ds-item-main">
                        <div class="ds-item-info">
                            <div class="ds-item-name">
                                <strong>${this.escapeHtml(ds.file_name)}</strong>
                            </div>
                            <div class="ds-item-meta">
                                <span>${ds.total_rows.toLocaleString()} 行</span>
                                <span>•</span>
                                <span>${ds.loaded_at}</span>
                            </div>
                        </div>
                        <div class="ds-item-actions">
                            <button class="btn btn-sm btn-danger delete-ds-btn" data-id="${ds.id}">
                                <span>🗑️</span> 删除
                            </button>
                        </div>
                    </div>
                </div>
            `;
        }).join('');
        
        console.log('数据源列表HTML已更新，绑定事件...');
        
        // 绑定删除事件
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
                } else if (feature === 'purchase') {
                    window.location.hash = 'purchase';
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
            // 支持多选文件
            const selected = await open({
                multiple: true,
                filters: [{
                    name: 'Excel文件',
                    extensions: ['xlsx', 'xls']
                }]
            });
            
            if (!selected) {
                console.log('用户取消了文件选择');
                return;
            }
            
            // 处理文件选择结果（可能是字符串或数组）
            const filePaths = Array.isArray(selected) ? selected : [selected];
            
            if (filePaths.length === 0) {
                return;
            }
            
            // 如果只有一个文件，直接导入
            if (filePaths.length === 1) {
                await this.importSingleFile(filePaths[0]);
                return;
            }
            
            // 多个文件，直接使用排队导入
            await this.importFilesSequential(filePaths);
            
        } catch (error) {
            console.error('选择文件失败:', error);
            if (error !== '用户取消操作') {
                this.showError('选择文件失败: ' + error);
            }
        }
    }
    
    async importSingleFile(filePath) {
        const { invoke } = window.__TAURI__.core;
        
        console.log('选择的文件:', filePath);
        
        // 显示导入进度提示
        this.showImportProgress(filePath, 1, 1);
        
        try {
            console.log('开始添加数据源...');
            const result = await invoke('add_data_source', { filePath: filePath });
            console.log('添加数据源结果:', result);
            
            // 隐藏进度提示
            this.hideProgress();
            
            this.showToast(`✅ 数据源添加成功！共 ${result.total_rows.toLocaleString()} 行数据`);
            
            // 立即重新加载数据源列表
            await this.loadDataSourceInfo();
            
            // 确保UI已更新，如果失败则重试
            setTimeout(async () => {
                await this.loadDataSourceInfo();
            }, 300);
        } catch (error) {
            console.error('添加数据源失败:', error);
            this.hideProgress();
            this.showError('添加数据源失败: ' + error);
        }
    }
    
    async importFilesSequential(filePaths) {
        const { invoke } = window.__TAURI__.core;
        const totalFiles = filePaths.length;
        let successCount = 0;
        let failCount = 0;
        const errors = [];
        
        // 显示多文件进度条
        if (window.app && window.app.showGlobalProgress) {
            window.app.showGlobalProgress('', 0, totalFiles, filePaths);
        }
        
        // 存储当前导入的文件索引，用于进度事件监听
        this.currentImportingFileIndex = null;
        
        try {
            // 串行导入文件
            for (let i = 0; i < filePaths.length; i++) {
                const filePath = filePaths[i];
                this.currentImportingFileIndex = i;
                
                // 更新文件状态为"导入中"
                if (window.app && window.app.updateFileProgress) {
                    window.app.updateFileProgress(i, 0, '导入中', '0%');
                }
                
                try {
                    const result = await invoke('add_data_source', { filePath: filePath });
                    successCount++;
                    
                    // 更新文件进度为完成
                    if (window.app && window.app.updateFileProgress) {
                        window.app.updateFileProgress(i, 100, '完成', '100%');
                    }
                    
                    // 每个文件导入成功后立即更新数据源列表
                    await this.loadDataSourceInfo();
                } catch (error) {
                    failCount++;
                    errors.push({ filePath, error: String(error) });
                    console.error(`导入文件失败:`, error);
                    
                    // 更新文件进度为失败
                    if (window.app && window.app.updateFileProgress) {
                        window.app.updateFileProgress(i, 100, '失败', '失败');
                    }
                }
            }
            
            this.currentImportingFileIndex = null;
            
            // 隐藏进度提示
            this.hideProgress();
            
            // 显示结果
            if (failCount === 0) {
                this.showToast(`✅ 成功导入 ${successCount} 个数据源！`);
            } else {
                this.showError(`导入完成：成功 ${successCount} 个，失败 ${failCount} 个\n\n失败文件：\n${errors.map(e => e.filePath.split(/[/\\]/).pop()).join('\n')}`);
            }
            
            // 最后再次刷新数据源列表，确保所有数据都已更新
            await this.loadDataSourceInfo();
            
        } catch (error) {
            console.error('批量导入失败:', error);
            this.currentImportingFileIndex = null;
            this.hideProgress();
            this.showError('批量导入失败: ' + error);
        }
    }
    
    showImportProgress(fileName = '', current = 1, total = 1, filePaths = []) {
        // 使用全局进度显示
        if (window.app && window.app.showGlobalProgress) {
            window.app.showGlobalProgress(fileName, current, total, filePaths);
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

