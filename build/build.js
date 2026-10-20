/*
 * src/index.html → index.html (학생에게 보이는 페이지)
 *  - 화면 코드(JSX)를 미리 변환하고 압축 → 브라우저에서 변환할 필요 없음, 소스가 읽기 어려워짐
 *  - Tailwind 디자인을 미리 만들어 넣음 → 외부 디자인 스크립트 불필요
 *  - 설정(config.js)은 압축하지 않고 따로 불러옴
 * 사용법: npm install → npm run build
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const babel = require('@babel/core');
const { minify } = require('terser');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'src', 'index.html');
const OUT = path.join(ROOT, 'index.html');

(async () => {
  let html = fs.readFileSync(SRC, 'utf8');

  // 1) 디자인 설정(tailwind.config)을 원본에서 그대로 읽기
  const cfgTag = html.match(/<script id="tailwind-config">([\s\S]*?)<\/script>/);
  if (!cfgTag) throw new Error('src/index.html에서 <script id="tailwind-config">를 찾지 못했습니다.');
  const sandbox = { tailwind: {} };
  vm.runInNewContext(cfgTag[1], sandbox);
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'hk-build-'));
  const twCfg = path.join(tmp, 'tailwind.config.js');
  fs.writeFileSync(twCfg, `module.exports = ${JSON.stringify({ content: [SRC], theme: sandbox.tailwind.config.theme || {} })};`);
  fs.writeFileSync(path.join(tmp, 'in.css'), '@tailwind base;\n@tailwind components;\n@tailwind utilities;\n');
  const twBin = path.join(ROOT, 'node_modules', '.bin', process.platform === 'win32' ? 'tailwindcss.cmd' : 'tailwindcss');
  execFileSync(twBin, ['-c', twCfg, '-i', path.join(tmp, 'in.css'), '-o', path.join(tmp, 'out.css'), '--minify'], { stdio: 'pipe' });
  const css = fs.readFileSync(path.join(tmp, 'out.css'), 'utf8');

  // 2) 화면 코드: JSX 변환 → 압축
  const appTag = html.match(/<script type="text\/babel">([\s\S]*?)<\/script>/);
  if (!appTag) throw new Error('src/index.html에서 <script type="text/babel">를 찾지 못했습니다.');
  const { code } = await babel.transformAsync(appTag[1], { presets: ['@babel/preset-react'], babelrc: false, configFile: false, compact: true });
  const min = await minify(code, { compress: { passes: 2 }, mangle: { toplevel: true }, format: { comments: false } });

  // 3) 결과 페이지 만들기
  html = html
    .replace(/\s*<!-- Tailwind CSS CDN -->\s*<script src="https:\/\/cdn\.tailwindcss\.com[^"]*"><\/script>/, '')
    .replace(/\s*<script src="https:\/\/unpkg\.com\/@babel\/standalone[^"]*"><\/script>/, '')
    .replace(cfgTag[0], `<style>${css}</style>`)
    .replace('<script src="../config.js"></script>', '<script src="config.js"></script>')
    .replace(appTag[0], () => `<script>${min.code}</script>`)
    .replace(/<!--(?!\s*\[if)[\s\S]*?-->/g, '')
    .replace(/\n\s*\n+/g, '\n');
  html = html.replace('<head>', '<head>\n<!-- 자동으로 만들어진 파일입니다. 수정은 src/index.html(화면) 또는 config.js(설정)에서 하세요. -->');
  fs.writeFileSync(OUT, html);
  fs.rmSync(tmp, { recursive: true, force: true });
  const kb = (n) => (n / 1024).toFixed(0) + 'KB';
  console.log(`완료: index.html ${kb(Buffer.byteLength(html))} (원본 ${kb(fs.statSync(SRC).size)}, CSS ${kb(css.length)}, 코드 ${kb(min.code.length)})`);
})().catch(e => { console.error('빌드 실패:', e.message); process.exit(1); });
