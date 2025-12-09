import { defineConfig } from 'vite';

export default defineConfig({
    // 开发服务器配置
    server: {
        port: 5173,
        strictPort: true,
    },
    
    // 构建配置
    build: {
        target: 'esnext',
        outDir: 'dist',
        emptyOutDir: true,
        // 内联小于4kb的资源
        assetsInlineLimit: 4096,
    },
    
    // 清除控制台
    clearScreen: false,
});

