import sharp from 'sharp';

// 创建一个简单的512x512蓝色渐变图标
const size = 512;

// 创建SVG
const svg = `
<svg width="${size}" height="${size}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="grad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#3b82f6;stop-opacity:1" />
      <stop offset="100%" style="stop-color:#8b5cf6;stop-opacity:1" />
    </linearGradient>
  </defs>
  <rect width="${size}" height="${size}" rx="100" fill="url(#grad)"/>
  <text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" 
        font-family="Arial, sans-serif" font-size="200" font-weight="bold" fill="white">
    CA
  </text>
</svg>
`;

await sharp(Buffer.from(svg))
    .png()
    .toFile('app-icon.png');

console.log('✅ app-icon.png created');

