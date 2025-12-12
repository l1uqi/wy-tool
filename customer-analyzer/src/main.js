// 主入口文件 - 路由和页面管理
import { HomePage } from './pages/home.js';
import { Top20Page } from './pages/top20.js';
import { MonthlyPage } from './pages/monthly.js';
import { RebatePage } from './pages/rebate.js';
import { PurchasePage } from './pages/purchase.js';

class App {
    constructor() {
        this.currentPage = 'home';
        this.pages = {
            home: new HomePage(this),
            top20: new Top20Page(this),
            monthly: new MonthlyPage(this),
            rebate: new RebatePage(this),
            purchase: new PurchasePage(this)
        };
        this.unlistenProgress = null;
        this.isImporting = false;
        
        this.init();
    }
    
    init() {
        this.render();
        this.bindNavEvents();
        this.setupGlobalProgressListener();
        
        // 检查URL hash进行初始导航
        const hash = window.location.hash.slice(1) || 'home';
        this.navigateTo(hash);
        
        // 监听hash变化
        window.addEventListener('hashchange', () => {
            const hash = window.location.hash.slice(1) || 'home';
            this.navigateTo(hash);
        });
    }
    
    async setupGlobalProgressListener() {
        if (!window.__TAURI__) return;
        
        const { listen } = window.__TAURI__.event;
        
        if (this.unlistenProgress) {
            this.unlistenProgress();
        }
        
        this.unlistenProgress = await listen('excel-progress', (event) => {
            const progress = event.payload;
            this.updateGlobalProgress(progress);
        });
    }
    
    showGlobalProgress(fileName = '', current = 1, total = 1, filePaths = []) {
        this.isImporting = true;
        const progressContainer = document.getElementById('globalImportProgress');
        const closeBtn = document.getElementById('globalProgressClose');
        
        if (!progressContainer) return;
        
        progressContainer.style.display = 'block';
        
        if (total > 1 && filePaths.length > 0) {
            // 批量导入：为每个文件创建独立的进度条
            this.renderMultiFileProgress(filePaths);
        } else {
            // 单个文件导入：使用原有单进度条
            this.renderSingleFileProgress(fileName);
        }
        
        if (closeBtn) {
            closeBtn.style.display = 'none';
        }
    }
    
    renderSingleFileProgress(fileName) {
        const progressContainer = document.getElementById('globalImportProgress');
        const content = progressContainer.querySelector('.global-progress-content');
        
        if (content) {
            content.innerHTML = `
                <div class="global-progress-header">
                    <h4>📥 正在导入数据源</h4>
                </div>
                <div class="progress-bar-container">
                    <div class="progress-bar" id="globalProgressBar"></div>
                </div>
                <p class="progress-text" id="globalProgressText">${fileName ? `正在导入: ${fileName}` : '正在导入数据源，请稍候...'}</p>
                <p class="progress-detail" id="globalProgressDetail"></p>
            `;
        }
    }
    
    renderMultiFileProgress(filePaths) {
        const progressContainer = document.getElementById('globalImportProgress');
        const content = progressContainer.querySelector('.global-progress-content');
        
        if (!content) return;
        
        const filesHtml = filePaths.map((filePath, index) => {
            const fileName = filePath.split(/[/\\]/).pop() || filePath;
            return `
                <div class="file-progress-item" data-file-index="${index}" data-file-path="${this.escapeHtml(filePath)}">
                    <div class="file-progress-header">
                        <span class="file-progress-name">${this.escapeHtml(fileName)}</span>
                        <span class="file-progress-status" id="fileStatus_${index}">等待中</span>
                    </div>
                    <div class="progress-bar-container">
                        <div class="progress-bar file-progress-bar" id="fileProgressBar_${index}" style="width: 0%"></div>
                    </div>
                    <p class="file-progress-text" id="fileProgressText_${index}">0%</p>
                </div>
            `;
        }).join('');
        
        content.innerHTML = `
            <div class="global-progress-header">
                <h4>📥 正在导入 ${filePaths.length} 个文件</h4>
            </div>
            <div class="multi-file-progress-list">
                ${filesHtml}
            </div>
        `;
    }
    
    updateFileProgress(fileIndex, percent, status = '导入中', text = '') {
        const progressBar = document.getElementById(`fileProgressBar_${fileIndex}`);
        const statusEl = document.getElementById(`fileStatus_${fileIndex}`);
        const textEl = document.getElementById(`fileProgressText_${fileIndex}`);
        
        if (progressBar) {
            progressBar.style.width = `${percent}%`;
        }
        
        if (statusEl) {
            statusEl.textContent = status;
            statusEl.className = `file-progress-status status-${status === '完成' ? 'success' : status === '失败' ? 'error' : 'pending'}`;
        }
        
        if (textEl) {
            textEl.textContent = text || `${percent}%`;
        }
    }
    
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    updateGlobalProgress(progress) {
        const progressBar = document.getElementById('globalProgressBar');
        const progressText = document.getElementById('globalProgressText');
        
        if (progressBar && progressText) {
            const percent = Math.round(progress.percent || 0);
            progressBar.style.width = `${percent}%`;
            progressText.textContent = `正在导入数据源... ${percent}%`;
        }
    }
    
    updateGlobalProgressOverall(percent, text) {
        const progressBar = document.getElementById('globalProgressBar');
        const progressText = document.getElementById('globalProgressText');
        
        if (progressBar && progressText) {
            progressBar.style.width = `${percent}%`;
            progressText.textContent = text || `正在导入数据源... ${percent}%`;
        }
    }
    
    updateGlobalProgressDetail(detail) {
        const progressDetail = document.getElementById('globalProgressDetail');
        if (progressDetail) {
            progressDetail.textContent = detail;
        }
    }
    
    hideGlobalProgress() {
        this.isImporting = false;
        const progressContainer = document.getElementById('globalImportProgress');
        const progressDetail = document.getElementById('globalProgressDetail');
        if (progressContainer) {
            progressContainer.style.display = 'none';
        }
        if (progressDetail) {
            progressDetail.textContent = '';
        }
    }
    
    render() {
        const app = document.getElementById('app');
        app.innerHTML = `
            <nav class="navbar">
                <a class="navbar-brand" data-page="home">
                    <div class="logo">📊</div>
                    <span>婉怡的工具箱</span>
                </a>
                <div class="navbar-nav">
                    <a class="nav-link" data-page="home">首页</a>
                    <a class="nav-link" data-page="top20">前20大客户</a>
                    <a class="nav-link" data-page="monthly">月度分析</a>
                    <a class="nav-link" data-page="rebate">返利分析</a>
                    <a class="nav-link" data-page="purchase">采购额计算</a>
                </div>
            </nav>
            <main id="page-content"></main>
        `;
        
        // 将全局进度显示附加到 body，避免页面切换时丢失
        this.renderGlobalProgress();
    }
    
    renderGlobalProgress() {
        // 如果已经存在，不重复创建
        if (document.getElementById('globalImportProgress')) {
            return;
        }
        
        const style = document.createElement('style');
        style.id = 'global-progress-styles';
        style.textContent = `
            .global-import-progress {
                position: fixed;
                bottom: 20px;
                right: 20px;
                background: var(--bg-card);
                border: 1px solid var(--border-color);
                border-radius: 12px;
                padding: 20px;
                min-width: 360px;
                max-width: 500px;
                max-height: 80vh;
                box-shadow: var(--shadow-lg);
                z-index: 10000;
                animation: slideInUp 0.3s ease;
                overflow-y: auto;
            }
            
            @keyframes slideInUp {
                from {
                    opacity: 0;
                    transform: translateY(20px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }
            
            .global-progress-content {
                width: 100%;
            }
            
            .global-progress-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 12px;
            }
            
            .global-progress-header h4 {
                margin: 0;
                color: var(--text-primary);
                font-size: 1rem;
                font-weight: 600;
            }
            
            .global-progress-close {
                background: transparent;
                border: none;
                color: var(--text-secondary);
                font-size: 1.2rem;
                cursor: pointer;
                padding: 4px 8px;
                border-radius: 4px;
                transition: all 0.2s;
            }
            
            .global-progress-close:hover {
                background: var(--bg-hover);
                color: var(--text-primary);
            }
            
            .progress-bar-container {
                width: 100%;
                height: 8px;
                background: var(--bg-secondary);
                border-radius: 4px;
                overflow: hidden;
                margin-bottom: 8px;
            }
            
            .progress-bar {
                height: 100%;
                background: linear-gradient(90deg, var(--accent-blue), var(--accent-purple));
                border-radius: 4px;
                transition: width 0.3s ease;
                width: 0%;
            }
            
            .progress-text {
                margin: 0;
                color: var(--text-primary);
                font-size: 0.9rem;
                font-weight: 500;
            }
            
            .progress-detail {
                margin: 0;
                color: var(--text-muted);
                font-size: 0.85rem;
                word-break: break-all;
            }
            
            .multi-file-progress-list {
                display: flex;
                flex-direction: column;
                gap: 16px;
                max-height: 400px;
                overflow-y: auto;
                padding-right: 4px;
            }
            
            .file-progress-item {
                padding: 12px;
                background: var(--bg-secondary);
                border-radius: 8px;
                border: 1px solid var(--border-color);
            }
            
            .file-progress-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 8px;
            }
            
            .file-progress-name {
                font-size: 0.9rem;
                color: var(--text-primary);
                font-weight: 500;
                flex: 1;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                margin-right: 8px;
            }
            
            .file-progress-status {
                font-size: 0.8rem;
                padding: 2px 8px;
                border-radius: 4px;
                font-weight: 500;
                white-space: nowrap;
            }
            
            .file-progress-status.status-pending {
                color: var(--text-muted);
                background: var(--bg-secondary);
            }
            
            .file-progress-status.status-success {
                color: #10b981;
                background: rgba(16, 185, 129, 0.1);
            }
            
            .file-progress-status.status-error {
                color: #ef4444;
                background: rgba(239, 68, 68, 0.1);
            }
            
            .file-progress-bar {
                background: linear-gradient(90deg, var(--accent-blue), var(--accent-purple));
            }
            
            .file-progress-text {
                margin: 4px 0 0 0;
                font-size: 0.8rem;
                color: var(--text-secondary);
            }
        `;
        document.head.appendChild(style);
        
        const progressDiv = document.createElement('div');
        progressDiv.id = 'globalImportProgress';
        progressDiv.className = 'global-import-progress';
        progressDiv.style.display = 'none';
        progressDiv.innerHTML = `
            <div class="global-progress-content">
                <div class="global-progress-header">
                    <h4>📥 正在导入数据源</h4>
                    <button class="global-progress-close" id="globalProgressClose" style="display: none;">✕</button>
                </div>
                <div class="progress-bar-container">
                    <div class="progress-bar" id="globalProgressBar"></div>
                </div>
                <p class="progress-text" id="globalProgressText">正在导入数据源...</p>
                <p class="progress-detail" id="globalProgressDetail" style="margin-top: 8px; font-size: 0.9rem; color: var(--text-muted);"></p>
            </div>
        `;
        document.body.appendChild(progressDiv);
    }
    
    bindNavEvents() {
        document.querySelectorAll('[data-page]').forEach(el => {
            el.addEventListener('click', (e) => {
                e.preventDefault();
                const page = el.dataset.page;
                window.location.hash = page;
            });
        });
    }
    
    navigateTo(pageName) {
        if (!this.pages[pageName]) {
            pageName = 'home';
        }
        
        this.currentPage = pageName;
        
        // 更新导航高亮
        document.querySelectorAll('.nav-link').forEach(el => {
            el.classList.toggle('active', el.dataset.page === pageName);
        });
        
        // 渲染页面
        const content = document.getElementById('page-content');
        content.innerHTML = '';
        this.pages[pageName].render(content);
    }
}

// 启动应用
document.addEventListener('DOMContentLoaded', () => {
    window.app = new App();
});
