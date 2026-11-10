// 将 Tabler 线性图标（MIT）按指定颜色/线宽栅格化为 PNG，供 build_deck.py 插入 PPT。
// 用法：node render_icons.js specs.json
//   specs.json: [{"name":"database","color":"0A4CFF","out":"assets/icons/database-0A4CFF.png","stroke":1.6,"px":256}, ...]
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

// @tabler/icons 的 exports 不暴露 package.json，这里沿 node 模块搜索路径查找 icons 目录
const searchDirs = [...(require.resolve.paths('@tabler/icons') || []), ...(process.env.NODE_PATH || '').split(path.delimiter)];
const iconDir = searchDirs.map((d) => path.join(d, '@tabler', 'icons', 'icons')).find((d) => d && fs.existsSync(d));
if (!iconDir) throw new Error('找不到 @tabler/icons，请先在 tools/ 下执行 npm install');

async function main() {
  const specs = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
  for (const s of specs) {
    const kind = s.filled ? 'filled' : 'outline';
    let svg = fs.readFileSync(path.join(iconDir, kind, `${s.name}.svg`), 'utf8');
    svg = svg.replace(/currentColor/g, `#${s.color}`)
             .replace(/stroke-width="2"/, `stroke-width="${s.stroke || 1.6}"`)
             .replace(/width="24"/, `width="${s.px || 256}"`)
             .replace(/height="24"/, `height="${s.px || 256}"`);
    fs.mkdirSync(path.dirname(s.out), { recursive: true });
    await sharp(Buffer.from(svg)).resize(s.px || 256, s.px || 256).png().toFile(s.out);
  }
}
main().catch((e) => { console.error(e); process.exit(1); });
