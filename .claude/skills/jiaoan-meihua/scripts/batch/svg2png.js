// 需要 sharp：npm install sharp（仓库 tools/ 已装）；用法 node svg2png.js <目录>（m_ 开头的缩略图底色 = 卡片底色）
const sharp = require(process.env.SHARP_PATH || 'sharp');
const fs = require('fs'), path = require('path');
const dir = process.argv[2];
(async () => {
  for (const f of fs.readdirSync(dir).filter(f => f.endsWith('.svg'))) {
    await sharp(fs.readFileSync(path.join(dir, f)), { density: 144 }).flatten({ background: f.startsWith('m_') ? '#F7F9FC' : '#ffffff' }).png().toFile(path.join(dir, f.replace('.svg', '.png')));
  }
  console.log('done');
})();
