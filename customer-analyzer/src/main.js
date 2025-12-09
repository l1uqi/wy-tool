// 主入口文件 - 路由和页面管理
import { HomePage } from './pages/home.js';
import { Top20Page } from './pages/top20.js';
import { MonthlyPage } from './pages/monthly.js';

class App {
    constructor() {
        this.currentPage = 'home';
        this.pages = {
            home: new HomePage(this),
            top20: new Top20Page(this),
            monthly: new MonthlyPage(this)
        };
        
        this.init();
    }
    
    init() {
        this.render();
        this.bindNavEvents();
        
        // 检查URL hash进行初始导航
        const hash = window.location.hash.slice(1) || 'home';
        this.navigateTo(hash);
        
        // 监听hash变化
        window.addEventListener('hashchange', () => {
            const hash = window.location.hash.slice(1) || 'home';
            this.navigateTo(hash);
        });
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
                </div>
            </nav>
            <main id="page-content"></main>
        `;
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
