// 首页组件
export class HomePage {
    constructor(app) {
        this.app = app;
    }
    
    render(container) {
        container.innerHTML = `
            <div class="home-container">
                <div class="home-hero slide-up">
                    <h1 class="home-title">客户分析工具</h1>
                    <p class="home-subtitle">
                        高性能数据分析工具，使用 Rust 后端处理大型 Excel 文件，
                        速度提升 10 倍以上
                    </p>
                </div>
                
                <div class="features-grid">
                    <div class="feature-card slide-up" data-feature="top20">
                        <div class="feature-icon blue">📊</div>
                        <h3 class="feature-title">前20大客户分析</h3>
                        <p class="feature-desc">
                            快速分析 Excel 数据，自动汇总客户订单金额，
                            生成前20大客户排行榜，支持百万级数据秒级处理
                        </p>
                        <span class="feature-badge rust">⚡ Rust高性能</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="monthly">
                        <div class="feature-icon purple">📈</div>
                        <h3 class="feature-title">月度销售分析</h3>
                        <p class="feature-desc">
                            按客户或地区统计每月销售额，生成可视化趋势图表，
                            支持环比增长分析
                        </p>
                        <span class="feature-badge new">✨ 新功能</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon cyan">🗂️</div>
                        <h3 class="feature-title">客户分类管理</h3>
                        <p class="feature-desc">
                            基于消费金额和频次自动将客户分为 VIP、普通、
                            潜力等类别
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon green">📋</div>
                        <h3 class="feature-title">报表生成</h3>
                        <p class="feature-desc">
                            一键生成专业的数据分析报表，支持多种格式导出
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon amber">🔄</div>
                        <h3 class="feature-title">数据对比</h3>
                        <p class="feature-desc">
                            多文件数据对比分析，快速发现数据差异和变化
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                    
                    <div class="feature-card slide-up" data-feature="coming">
                        <div class="feature-icon rose">⚙️</div>
                        <h3 class="feature-title">批量处理</h3>
                        <p class="feature-desc">
                            批量导入多个 Excel 文件，自动合并处理，
                            提高工作效率
                        </p>
                        <span class="feature-badge coming">🚧 开发中</span>
                    </div>
                </div>
            </div>
        `;
        
        this.bindEvents(container);
    }
    
    bindEvents(container) {
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
