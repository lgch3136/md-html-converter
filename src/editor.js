// === MD ↔ HTML 编辑器 ===
import { marked } from 'marked';
import TurndownService from 'turndown';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, ExternalHyperlink } from 'docx';
import mammoth from 'mammoth';

// --- 初始化库 ---
const turndown = new TurndownService({
  headingStyle: 'atx', codeBlockStyle: 'fenced', hr: '---',
  bulletListMarker: '-', emDelimiter: '*', strongDelimiter: '**',
});
marked.setOptions({ breaks: true, gfm: true, headerIds: false, mangle: false });

// ===== DOM =====
const mdRender   = document.getElementById('mdRenderPane');
const mdSource   = document.getElementById('mdSourcePane');
const htmlRender = document.getElementById('htmlRenderPane');
const htmlSource = document.getElementById('htmlSourcePane');
const mdToolbar  = document.getElementById('mdToolbar');
const htmlToolbar = document.getElementById('htmlToolbar');
const fileInfo   = document.getElementById('fileInfo');
const btnClear   = document.getElementById('btnClear');

const btnSyncMdToHtml = document.getElementById('btnSyncMdToHtml');
const btnSyncHtmlToMd = document.getElementById('btnSyncHtmlToMd');

const importBtn    = document.getElementById('btnImport');
const importMenu   = document.getElementById('importMenu');
const importInput  = document.getElementById('importInput');

const exportBtn    = document.getElementById('btnExport');
const exportMenu   = document.getElementById('exportMenu');

const btnTemplate  = document.getElementById('btnTemplate');
const templateMenu = document.getElementById('templateMenu');

const videoModal      = document.getElementById('videoModal');
const videoUrl        = document.getElementById('videoUrl');
const tableModal      = document.getElementById('tableModal');
const tableRows       = document.getElementById('tableRows');
const tableCols       = document.getElementById('tableCols');
const codeBlockModal  = document.getElementById('codeBlockModal');
const codeLang        = document.getElementById('codeLang');
const codeContent     = document.getElementById('codeContent');
const linkModal       = document.getElementById('linkModal');
const linkText        = document.getElementById('linkText');
const linkUrl         = document.getElementById('linkUrl');
const imageModal      = document.getElementById('imageModal');
const imgFileInput    = document.getElementById('imgFileInput');

const imageInputMd    = document.getElementById('imageInputMd');
const imageInputHtml  = document.getElementById('imageInputHtml');

// ===== 模板库（函数返回字符串，避免反引号冲突） =====
function tplBlog() {
  return '# 技术博客\n\n> 作者：你的名字\n> 日期：' + new Date().toLocaleDateString('zh-CN') + '\n\n## 概述\n\n在这里写下你的技术分享内容…\n\n## 核心内容\n\n1. 第一点\n2. 第二点\n3. 第三点\n\n## 代码示例\n\n```javascript\n// 你的代码\nconsole.log("Hello World");\n```\n\n## 总结\n\n总结你的观点和收获。\n\n## 参考资料\n\n- [链接1](https://example.com)\n- [链接2](https://example.com)';
}

function tplProduct() {
  return '# 产品介绍\n\n> 版本：v1.0  |  更新日期：' + new Date().toLocaleDateString('zh-CN') + '\n\n## 产品概述\n\n一句话介绍你的产品是做什么的。\n\n## 核心功能\n\n| 功能 | 描述 | 状态 |\n|------|------|------|\n| 功能一 | 描述说明 | ✅ 已完成 |\n| 功能二 | 描述说明 | 🔨 开发中 |\n| 功能三 | 描述说明 | 📋 规划中 |\n\n## 使用场景\n\n1. 场景一描述\n2. 场景二描述\n3. 场景三描述\n\n## 快速开始\n\n```bash\n# 安装\nnpm install\n\n# 运行\nnpm start\n```\n\n## 联系我们\n\n- 邮箱：your@email.com\n- 官网：https://example.com';
}

function tplMeeting() {
  return '# 会议纪要\n\n> 日期：' + new Date().toLocaleDateString('zh-CN') + '\n> 参会人：张三、李四、王五\n> 主持人：张三\n\n## 会议主题\n\n简要描述本次会议的核心议题。\n\n## 讨论内容\n\n### 议题一\n\n- 背景说明\n- 各方意见\n- 决议结果\n\n### 议题二\n\n- 背景说明\n- 各方意见\n- 决议结果\n\n## Action Items\n\n| 事项 | 负责人 | 截止日期 | 状态 |\n|------|--------|----------|------|\n| 任务一 | 张三 | 2026-06-01 | ⬜ 待办 |\n| 任务二 | 李四 | 2026-06-05 | ⬜ 待办 |\n| 任务三 | 王五 | 2026-06-10 | ⬜ 待办 |\n\n## 下次会议\n\n时间：待定';
}

function tplResume() {
  return '# 个人简历\n\n> 姓名：你的姓名\n> 电话：138-xxxx-xxxx\n> 邮箱：your@email.com\n> 所在地：城市名\n\n## 求职意向\n\n目标岗位：高级前端工程师\n期望薪资：面议\n到岗时间：随时 / 1个月内\n\n## 教育背景\n\n### 学校名称 — 专业（本科/硕士）\n2018.09 - 2022.06\n\n- GPA：3.8/4.0\n- 相关课程：数据结构、算法、操作系统\n\n## 工作经历\n\n### 公司名称 — 前端工程师\n2022.07 - 至今\n\n- 负责 xxx 项目的核心页面开发与维护\n- 使用 React/Vue 技术栈，提升了 40% 的开发效率\n- 主导组件库建设，沉淀 50+ 可复用组件\n\n### 公司名称 — 前端实习生\n2021.06 - 2022.06\n\n- 参与 xxx 项目需求开发与迭代\n- 协助完成性能优化，首屏加载速度提升 30%\n\n## 项目经历\n\n### 项目名称\n- 项目描述：一句话介绍项目\n- 技术栈：React + TypeScript + Node.js\n- 个人职责：核心模块开发、性能优化\n\n## 技能清单\n\n- 编程语言：JavaScript / TypeScript / Python\n- 前端框架：React / Vue\n- 工具：Webpack / Vite / Git\n- 其他：熟练使用 Figma 设计协作';
}

function tplApi() {
  return '# API 文档\n\n> 版本：v1.0  |  Base URL：`https://api.example.com`\n> 最后更新：' + new Date().toLocaleDateString('zh-CN') + '\n\n## 认证方式\n\n所有 API 请求需在 Header 中携带 Token：\n\n```\nAuthorization: Bearer <your-token>\n```\n\n## 接口列表\n\n### 获取用户信息\n\n- **URL**：`GET /api/users/:id`\n- **描述**：根据用户 ID 获取用户详细信息\n- **请求参数**：\n\n| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n| id | string | 是 | 用户唯一标识 |\n\n- **响应示例**：\n\n```json\n{\n  "code": 200,\n  "data": {\n    "id": "123",\n    "name": "张三",\n    "email": "zhangsan@example.com"\n  }\n}\n```\n\n### 创建用户\n\n- **URL**：`POST /api/users`\n- **描述**：创建新用户\n- **请求体**：\n\n```json\n{\n  "name": "新用户",\n  "email": "new@example.com"\n}\n```\n\n- **响应状态码**：\n  - `201`：创建成功\n  - `400`：参数错误\n  - `409`：用户已存在\n\n## 错误码\n\n| 状态码 | 含义 |\n|--------|------|\n| 200 | 成功 |\n| 400 | 请求参数错误 |\n| 401 | 未授权 |\n| 404 | 资源不存在 |\n| 500 | 服务器内部错误 |';
}

function tplNotes() {
  return '# 读书笔记：《书名》\n\n> 作者：作者名\n> 阅读日期：' + new Date().toLocaleDateString('zh-CN') + '\n> 评分：★★★★☆（5 星满分）\n\n## 📖 书籍概述\n\n用 2-3 句话概括这本书的核心内容和主旨。\n\n## 📝 核心观点\n\n### 观点一：标题\n\n用自己的话总结这个观点，标注原文页码。\n\n> "原文引用放在这里" —— 作者名\n\n### 观点二：标题\n\n用自己的话总结这个观点。\n\n### 观点三：标题\n\n用自己的话总结这个观点。\n\n## 💡 感悟与启发\n\n1. 这本书让我意识到……\n2. 可以在实际中应用的点：\n   - \n   - \n\n## 📌 金句摘录\n\n- "金句一"\n- "金句二"\n- "金句三"\n\n## 🔗 相关推荐\n\n- [相关文章](https://example.com)\n- [作者其他作品]()';
}

const templateFns = { blog: tplBlog, product: tplProduct, meeting: tplMeeting, resume: tplResume, api: tplApi, notes: tplNotes };

// ===== 状态 =====
const paneModes = { md: 'render', html: 'render' };
let pendingTarget = null;

// ========================
//  面板模式切换
// ========================
document.querySelectorAll('.mdhtml-mode-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    setPaneMode(btn.dataset.pane, btn.dataset.mode);
  });
});

function setPaneMode(pane, mode, skipContentSync = false) {
  const prevMode = paneModes[pane];
  paneModes[pane] = mode;

  document.querySelectorAll('.mdhtml-mode-btn[data-pane="' + pane + '"]').forEach(b => {
    b.classList.toggle('active', b.dataset.mode === mode);
  });

  if (pane === 'md') {
    if (mode === 'render') {
      if (!skipContentSync && prevMode === 'source') {
        mdRender.innerHTML = mdRenderFromMd(mdSource.value);
      }
      mdSource.classList.add('hidden');
      mdToolbar.classList.remove('hidden');
      mdRender.classList.remove('hidden');
    } else {
      if (!skipContentSync && prevMode === 'render') {
        mdSource.value = extractMdFromRender(mdRender);
      }
      mdRender.classList.add('hidden');
      mdToolbar.classList.add('hidden');
      mdSource.classList.remove('hidden');
      mdSource.focus();
    }
  } else {
    if (mode === 'render') {
      if (!skipContentSync && prevMode === 'source') {
        htmlRender.innerHTML = htmlSource.value;
      }
      htmlSource.classList.add('hidden');
      htmlToolbar.classList.remove('hidden');
      htmlRender.classList.remove('hidden');
    } else {
      if (!skipContentSync && prevMode === 'render') {
        htmlSource.value = htmlRender.innerHTML;
      }
      htmlRender.classList.add('hidden');
      htmlToolbar.classList.add('hidden');
      htmlSource.classList.remove('hidden');
      htmlSource.focus();
    }
  }
  updateInfo();
}

// ===== MD → render HTML =====
function mdRenderFromMd(md) {
  if (!md.trim()) return '';
  return marked.parse(md);
}

// ===== contenteditable → Markdown =====
function extractMdFromRender(el) {
  return turndown.turndown(el.innerHTML || '');
}

// ========================
//  内容变化 ↔ 实时预览
// ========================
mdSource.addEventListener('input', () => {
  if (paneModes.md === 'render') {
    mdRender.innerHTML = mdRenderFromMd(mdSource.value);
  }
  updateInfo();
});

htmlSource.addEventListener('input', () => {
  if (paneModes.html === 'render') {
    htmlRender.innerHTML = htmlSource.value;
  }
  updateInfo();
});

mdRender.addEventListener('input', () => updateInfo());
htmlRender.addEventListener('input', () => updateInfo());

// ========================
//  互转按钮
// ========================

btnSyncMdToHtml.addEventListener('click', () => {
  let md;
  if (paneModes.md === 'render') {
    md = extractMdFromRender(mdRender);
  } else {
    md = mdSource.value;
  }
  md = md.trim();
  if (!md) { toast('⚠️ 请先在左侧输入 Markdown 内容'); return; }

  mdSource.value = md;
  const html = marked.parse(md);

  htmlSource.value = html;
  setPaneMode('md', 'render', true);
  mdRender.innerHTML = html;
  setPaneMode('html', 'render', true);
  htmlRender.innerHTML = html;

  updateInfo();
  toast('MD → HTML ✅ 右侧为预览效果');
});

btnSyncHtmlToMd.addEventListener('click', () => {
  let html;
  if (paneModes.html === 'render') {
    html = htmlRender.innerHTML;
  } else {
    html = htmlSource.value;
  }
  html = html.trim();
  if (!html) { toast('⚠️ 请先在右侧输入 HTML 内容'); return; }

  htmlSource.value = html;
  const md = turndown.turndown(html);

  mdSource.value = md;
  setPaneMode('md', 'render', true);
  mdRender.innerHTML = html;
  setPaneMode('html', 'render', true);
  htmlRender.innerHTML = html;

  updateInfo();
  toast('HTML → MD ✅');
});

// ========================
//  富文本格式化
// ========================
function setupToolbar(toolbarEl, renderEl) {
  toolbarEl.querySelectorAll('.mdhtml-fmt-btn').forEach(btn => {
    btn.addEventListener('click', () => handleFormat(btn.dataset.fmt, renderEl));
  });
}
setupToolbar(mdToolbar, mdRender);
setupToolbar(htmlToolbar, htmlRender);

function handleFormat(fmt, el) {
  el.focus();
  switch (fmt) {
    case 'bold':       document.execCommand('bold'); break;
    case 'italic':     document.execCommand('italic'); break;
    case 'underline':  document.execCommand('underline'); break;
    case 'strike':     document.execCommand('strikeThrough'); break;
    case 'h1':         wrapBlock(el, '<h1>', '</h1>'); break;
    case 'h2':         wrapBlock(el, '<h2>', '</h2>'); break;
    case 'h3':         wrapBlock(el, '<h3>', '</h3>'); break;
    case 'ul':         wrapBlock(el, '<ul>\n  <li>', '</li>\n</ul>'); break;
    case 'ol':         wrapBlock(el, '<ol>\n  <li>', '</li>\n</ol>'); break;
    case 'checklist':  wrapBlock(el, '<ul style="list-style:none;padding-left:0;">\n  <li>☐ ', '</li>\n</ul>'); break;
    case 'blockquote': wrapBlock(el, '<blockquote>', '</blockquote>'); break;
    case 'hr':         document.execCommand('insertHTML', false, '<hr>'); break;
    case 'codeInline': {
      const sel = window.getSelection();
      if (sel && sel.toString()) document.execCommand('insertHTML', false, '<code>' + sel.toString() + '</code>');
      break;
    }
    case 'insertImage': openImageModal(el); break;
    case 'insertVideo': openVideoModal(el); break;
    case 'insertTable': openTableModal(el); break;
    case 'insertCodeBlock': openCodeBlockModal(el); break;
    case 'link':       openLinkModal(el); break;
    case 'undo':       document.execCommand('undo'); break;
    case 'redo':       document.execCommand('redo'); break;
  }
}

function wrapBlock(el, before, after) {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return;
  const range = sel.getRangeAt(0);
  const text  = range.toString() || '输入内容';
  const wrapper = document.createElement('div');
  wrapper.innerHTML = before + text + after;
  range.deleteContents();
  const frag = document.createDocumentFragment();
  while (wrapper.firstChild) frag.appendChild(wrapper.firstChild);
  range.insertNode(frag);
}

// ========================
//  弹窗管理
// ========================
function openModal(modal) { modal.classList.remove('hidden'); }
function closeModal(modal) { modal.classList.add('hidden'); }
function closeAllModals() {
  [videoModal, tableModal, codeBlockModal, linkModal, imageModal].forEach(closeModal);
  pendingTarget = null;
}

document.querySelectorAll('.mdhtml-modal-overlay').forEach(overlay => {
  overlay.addEventListener('click', e => { if (e.target === overlay) closeModal(overlay); });
});

document.getElementById('btnCancelVideo').onclick     = () => closeModal(videoModal);
document.getElementById('btnCancelTable').onclick     = () => closeModal(tableModal);
document.getElementById('btnCancelCodeBlock').onclick = () => closeModal(codeBlockModal);
document.getElementById('btnCancelLink').onclick      = () => closeModal(linkModal);
document.getElementById('btnCancelImage').onclick     = () => closeModal(imageModal);

function openImageModal(el)    { pendingTarget = el; imgFileInput.value = ''; openModal(imageModal); }
function openVideoModal(el)    { pendingTarget = el; videoUrl.value = ''; openModal(videoModal); videoUrl.focus(); }
function openTableModal(el)    { pendingTarget = el; tableRows.value = 3; tableCols.value = 3; openModal(tableModal); }
function openCodeBlockModal(el){ pendingTarget = el; codeLang.value = ''; codeContent.value = ''; openModal(codeBlockModal); codeContent.focus(); }
function openLinkModal(el)     { pendingTarget = el; const sel = window.getSelection(); linkText.value = (sel && sel.toString()) || ''; linkUrl.value = ''; openModal(linkModal); linkText.focus(); }

// ========================
//  弹窗确认
// ========================
document.getElementById('btnConfirmVideo').onclick = () => {
  const url = videoUrl.value.trim();
  if (!url || !pendingTarget) return;
  let src = url;
  if (url.includes('b23.tv') || /bilibili\.com\/BV/i.test(url)) {
    var m = url.match(/BV[\w]+/i);
    src = m ? 'https://player.bilibili.com/player.html?bvid=' + m[0] + '&autoplay=0' : url;
  } else if (url.includes('youtube.com') || url.includes('youtu.be')) {
    var m2 = url.match(/(?:v=|youtu\.be\/|embed\/)([\w-]{11})/i);
    src = m2 ? 'https://www.youtube.com/embed/' + m2[1] : url;
  }
  var iframeHtml = '<div style="position:relative;width:100%;padding-bottom:56.25%;height:0;overflow:hidden;margin:12px 0;"><iframe src="' + src + '" style="position:absolute;inset:0;width:100%;height:100%;border:none;border-radius:8px;" allowfullscreen></iframe></div>';
  pendingTarget.focus();
  document.execCommand('insertHTML', false, iframeHtml);
  closeModal(videoModal);
};

document.getElementById('btnConfirmTable').onclick = () => {
  var rows = parseInt(tableRows.value) || 3;
  var cols = parseInt(tableCols.value) || 3;
  if (!pendingTarget) return;
  if (pendingTarget === mdRender) {
    var md = '\n|';
    for (var c = 0; c < cols; c++) md += ' 列' + (c+1) + ' |';
    md += '\n|';
    for (var c2 = 0; c2 < cols; c2++) md += ' --- |';
    md += '\n';
    for (var r = 0; r < rows - 1; r++) { md += '|'; for (var c3 = 0; c3 < cols; c3++) md += ' 内容 |'; md += '\n'; }
    document.execCommand('insertHTML', false, marked.parse(md));
  } else {
    var t = '<table style="width:100%;border-collapse:collapse;margin:12px 0;">';
    t += '<thead><tr>';
    for (var c4 = 0; c4 < cols; c4++) t += '<th style="border:1px solid #21262d;padding:8px 12px;background:#161b22;color:#58a6ff;">列' + (c4+1) + '</th>';
    t += '</tr></thead><tbody>';
    for (var r2 = 0; r2 < rows - 1; r2++) { t += '<tr>'; for (var c5 = 0; c5 < cols; c5++) t += '<td style="border:1px solid #21262d;padding:8px 12px;">内容</td>'; t += '</tr>'; }
    document.execCommand('insertHTML', false, t + '</tbody></table>');
  }
  closeModal(tableModal);
};

document.getElementById('btnConfirmCodeBlock').onclick = () => {
  var lang = codeLang.value.trim();
  var code = codeContent.value;
  if (!pendingTarget) return;
  if (pendingTarget === mdRender) {
    var fence = lang ? '```' + lang + '\n' + code + '\n```' : '```\n' + code + '\n```';
    document.execCommand('insertHTML', false, marked.parse(fence));
  } else {
    document.execCommand('insertHTML', false, '<pre style="background:#161b22;border:1px solid #21262d;border-radius:6px;padding:12px 14px;overflow-x:auto;margin:12px 0;"><code>' + escapeHtml(code) + '</code></pre>');
  }
  closeModal(codeBlockModal);
};

document.getElementById('btnConfirmLink').onclick = () => {
  var text = linkText.value.trim() || '链接';
  var url  = linkUrl.value.trim();
  if (!url || !pendingTarget) return;
  pendingTarget.focus();
  document.execCommand('insertHTML', false, '<a href="' + url + '" target="_blank" rel="noopener" style="color:#79c0ff;">' + text + '</a>');
  closeModal(linkModal);
};

document.getElementById('btnConfirmImage').onclick = async () => {
  var file = imgFileInput.files[0];
  if (!file || !pendingTarget) return;
  var dataUrl = await fileToDataURL(file);
  var alt = file.name.replace(/\.[^.]+$/, '');
  pendingTarget.focus();
  if (pendingTarget === mdRender) {
    var mdImg = '![' + alt + '](data:image;base64,' + dataUrl.split(',')[1] + ')';
    document.execCommand('insertHTML', false, marked.parse(mdImg));
  } else {
    document.execCommand('insertHTML', false, '<img src="' + dataUrl + '" alt="' + alt + '" style="max-width:100%;border-radius:6px;margin:6px 0;display:block;" />');
  }
  closeModal(imageModal);
  fileInfo.textContent = '已插入: ' + file.name;
  toast('🖼️ ' + file.name);
};

// ========================
//  文件上传
// ========================
imageInputMd.addEventListener('change', async e => {
  var file = e.target.files[0];
  if (!file) return;
  await uploadFile(file, 'md');
  imageInputMd.value = '';
});

imageInputHtml.addEventListener('change', async e => {
  var file = e.target.files[0];
  if (!file) return;
  await uploadFile(file, 'html');
  imageInputHtml.value = '';
});

async function uploadFile(file, pane) {
  if (file.name.match(/\.(md|markdown|txt|html)$/i)) {
    var text = await file.text();
    if (pane === 'md') {
      mdSource.value = text;
      if (paneModes.md === 'render') mdRender.innerHTML = mdRenderFromMd(text);
    } else {
      htmlSource.value = text;
      if (paneModes.html === 'render') htmlRender.innerHTML = text;
    }
    fileInfo.textContent = '已加载: ' + file.name;
    toast('📄 ' + file.name);
  } else {
    var dataUrl = await fileToDataURL(file);
    var alt = file.name.replace(/\.[^.]+$/, '');
    var renderEl = pane === 'md' ? mdRender : htmlRender;
    renderEl.focus();
    if (pane === 'md') {
      var mdImg = '![' + alt + '](data:image;base64,' + dataUrl.split(',')[1] + ')';
      document.execCommand('insertHTML', false, marked.parse(mdImg));
    } else {
      document.execCommand('insertHTML', false, '<img src="' + dataUrl + '" alt="' + alt + '" style="max-width:100%;border-radius:6px;margin:6px 0;display:block;" />');
    }
    fileInfo.textContent = '已插入: ' + file.name;
    toast('🖼️ ' + file.name);
  }
  updateInfo();
}

// ========================
//  粘贴图片 (Ctrl+V)
// ========================
[mdRender, htmlRender].forEach(el => {
  el.addEventListener('paste', async e => {
    var items = (e.clipboardData || e.originalEvent && e.originalEvent.clipboardData) && e.clipboardData.items;
    if (!items) return;
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (item.type && item.type.startsWith('image/')) {
        e.preventDefault();
        var file = item.getAsFile();
        if (file) await uploadFile(file, el === mdRender ? 'md' : 'html');
      }
    }
  });
});

// ========================
//  拖拽文件
// ========================
function setupDropZone(el, pane) {
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag-over'); });
  el.addEventListener('dragleave', e => { if (!el.contains(e.relatedTarget)) el.classList.remove('drag-over'); });
  el.addEventListener('drop', async e => {
    e.preventDefault();
    el.classList.remove('drag-over');
    var file = e.dataTransfer.files[0];
    if (file) await uploadFile(file, pane);
  });
}
setupDropZone(mdRender, 'md');
setupDropZone(mdSource, 'md');
setupDropZone(htmlRender, 'html');
setupDropZone(htmlSource, 'html');

// ========================
//  键盘快捷键
// ========================
document.addEventListener('keydown', e => {
  if (e.ctrlKey || e.metaKey) {
    var pane = getActivePane();
    if (pane && paneModes[pane] === 'render') {
      if (e.key === 'b') { e.preventDefault(); document.execCommand('bold'); }
      if (e.key === 'i') { e.preventDefault(); document.execCommand('italic'); }
      if (e.key === 's') { e.preventDefault(); btnSyncMdToHtml.click(); }
      if (e.key === 'u') { e.preventDefault(); document.execCommand('underline'); }
    }
  }
});

function getActivePane() {
  if (document.activeElement === mdRender || document.activeElement === mdSource) return 'md';
  if (document.activeElement === htmlRender || document.activeElement === htmlSource) return 'html';
  return null;
}

// ========================
//  清空
// ========================
btnClear.addEventListener('click', () => {
  if (!confirm('确定清空所有内容？')) return;
  mdSource.value = ''; htmlSource.value = '';
  mdRender.innerHTML = ''; htmlRender.innerHTML = '';
  updateInfo();
  toast('已清空 🗑️');
});

// ========================
//  导入（统一入口）
// ========================
importBtn.addEventListener('click', e => {
  e.stopPropagation();
  importMenu.classList.toggle('hidden');
  exportMenu.classList.add('hidden');
  templateMenu.classList.add('hidden');
});

importMenu.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', e => {
    e.stopPropagation();
    var format = item.dataset.format;
    if (format === 'word') {
      importInput.accept = '.docx,.doc';
    } else if (format === 'md') {
      importInput.accept = '.md,.markdown,.txt';
    } else if (format === 'html') {
      importInput.accept = '.html,.htm';
    } else if (format === 'txt') {
      importInput.accept = '.txt';
    }
    importInput.dataset.importFormat = format;
    importInput.click();
    importMenu.classList.add('hidden');
  });
});

importInput.addEventListener('change', async e => {
  var file = e.target.files[0];
  if (!file) return;
  var format = importInput.dataset.importFormat || 'auto';
  await handleImport(file, format);
  importInput.value = '';
});

async function handleImport(file, format) {
  toast('正在导入…');
  try {
    if (format === 'word' || file.name.match(/\.(docx|doc)$/i)) {
      var arrayBuffer = await file.arrayBuffer();
      var result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
      var html = result.value;
      if (result.messages.length) console.warn('mammoth messages:', result.messages);
      var md = turndown.turndown(html);
      mdSource.value = md;
      htmlSource.value = html;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html;
      htmlRender.innerHTML = html;
      fileInfo.textContent = '已导入 Word: ' + file.name;
      toast('📥 已导入 Word: ' + file.name);
    } else if (format === 'md' || file.name.match(/\.(md|markdown)$/i)) {
      var text = await file.text();
      var html = marked.parse(text);
      mdSource.value = text;
      htmlSource.value = html;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html;
      htmlRender.innerHTML = html;
      fileInfo.textContent = '已导入 Markdown: ' + file.name;
      toast('📥 已导入 MD: ' + file.name);
    } else if (format === 'html' || file.name.match(/\.(html|htm)$/i)) {
      var html2 = await file.text();
      var md2 = turndown.turndown(html2);
      mdSource.value = md2;
      htmlSource.value = html2;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html2;
      htmlRender.innerHTML = html2;
      fileInfo.textContent = '已导入 HTML: ' + file.name;
      toast('📥 已导入 HTML: ' + file.name);
    } else if (format === 'txt' || file.name.match(/\.txt$/i)) {
      var text2 = await file.text();
      mdSource.value = text2;
      htmlSource.value = escapeHtml(text2).replace(/\n/g, '<br>');
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = '<p>' + escapeHtml(text2).replace(/\n/g, '<br>') + '</p>';
      htmlRender.innerHTML = '<p>' + escapeHtml(text2).replace(/\n/g, '<br>') + '</p>';
      fileInfo.textContent = '已导入文本: ' + file.name;
      toast('📥 已导入 TXT: ' + file.name);
    } else {
      try {
        var text3 = await file.text();
        var html3 = marked.parse(text3);
        mdSource.value = text3;
        htmlSource.value = html3;
        setPaneMode('md', 'render', true);
        setPaneMode('html', 'render', true);
        mdRender.innerHTML = html3;
        htmlRender.innerHTML = html3;
        fileInfo.textContent = '已导入: ' + file.name;
        toast('📥 已导入: ' + file.name);
      } catch (err) {
        toast('❌ 无法识别的文件格式');
        console.error(err);
      }
    }
    updateInfo();
  } catch (err) {
    console.error(err);
    toast('导入失败 ❌');
  }
}

// ========================
//  导出（统一入口）
// ========================
exportBtn.addEventListener('click', e => {
  e.stopPropagation();
  exportMenu.classList.toggle('hidden');
  importMenu.classList.add('hidden');
  templateMenu.classList.add('hidden');
});

exportMenu.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', e => {
    e.stopPropagation();
    exportMenu.classList.add('hidden');
    var format = item.dataset.format;
    handleExport(format);
  });
});

async function handleExport(format) {
  var md = mdSource.value.trim();
  if (!md) { toast('⚠️ 没有内容可导出'); return; }

  try {
    var htmlBody = marked.parse(md);

    switch (format) {
      case 'pdf':    await exportPdf(htmlBody); break;
      case 'word':   await exportWord(htmlBody); break;
      case 'html':   exportStandaloneHtml(htmlBody); break;
      case 'md':     exportMarkdown(md); break;
      case 'txt':    exportText(htmlBody); break;
      default:       toast('未知导出格式');
    }
  } catch (err) {
    console.error(err);
    toast(format.toUpperCase() + ' 导出失败 ❌');
  }
}

// --- PDF 导出 ---
async function exportPdf(htmlBody) {
  toast('正在生成 PDF…');
  var h2p = await loadHtml2Pdf();
  var fullHtml = buildFullHtml(htmlBody);

  var iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.left = '-9999px';
  iframe.style.width = '800px';
  iframe.style.height = '600px';
  document.body.appendChild(iframe);

  iframe.contentDocument.open();
  iframe.contentDocument.write(fullHtml);
  iframe.contentDocument.close();

  await new Promise(r => setTimeout(r, 500));

  var opt = {
    margin: [10, 10, 10, 10],
    filename: 'document.pdf',
    image: { type: 'jpeg', quality: 0.95 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  await h2p().set(opt).from(iframe.contentDocument.body).save();
  document.body.removeChild(iframe);
  toast('PDF 已导出 📄');
}

var html2pdfModule = null;
async function loadHtml2Pdf() {
  if (html2pdfModule) return html2pdfModule;
  if (window.html2pdf) return window.html2pdf;

  return new Promise(function(resolve, reject) {
    var s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.3/html2pdf.bundle.min.js';
    s.onload = function() {
      html2pdfModule = window.html2pdf;
      resolve(window.html2pdf);
    };
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// --- Word 导出 ---
async function exportWord(htmlBody) {
  toast('正在生成 Word…');
  var paragraphs = htmlToDocxParagraphs(htmlBody);
  var doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });
  var blob = await Packer.toBlob(doc);
  downloadBlob(blob, 'document.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  toast('Word 已导出 📤');
}

// --- 独立 HTML 导出 ---
function exportStandaloneHtml(htmlBody) {
  var full = buildFullHtml(htmlBody);
  downloadBlob(full, 'document.html', 'text/html;charset=utf-8');
  toast('HTML 已导出 🌐');
}

// --- MD 导出 ---
function exportMarkdown(md) {
  downloadBlob(md, 'document.md', 'text/markdown;charset=utf-8');
  toast('Markdown 已导出 📝');
}

// --- TXT 导出 ---
function exportText(htmlBody) {
  var div = document.createElement('div');
  div.innerHTML = htmlBody;
  var text = div.textContent || '';
  downloadBlob(text, 'document.txt', 'text/plain;charset=utf-8');
  toast('纯文本已导出 📃');
}

// ========================
//  模板功能
// ========================
btnTemplate.addEventListener('click', e => {
  e.stopPropagation();
  templateMenu.classList.toggle('hidden');
  importMenu.classList.add('hidden');
  exportMenu.classList.add('hidden');
});

templateMenu.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', e => {
    e.stopPropagation();
    var tpl = item.dataset.template;
    if (!templateFns[tpl]) return;
    templateMenu.classList.add('hidden');

    var md = mdSource.value.trim();
    var html = htmlSource.value.trim();
    if (md || html) {
      if (!confirm('当前内容将被覆盖，确定加载模板吗？')) return;
    }

    var templateContent = templateFns[tpl]();
    mdSource.value = templateContent;
    var renderedHtml = marked.parse(templateContent);
    htmlSource.value = renderedHtml;
    setPaneMode('md', 'render', true);
    setPaneMode('html', 'render', true);
    mdRender.innerHTML = renderedHtml;
    htmlRender.innerHTML = renderedHtml;
    updateInfo();
    toast('📋 模板已加载');
  });
});

// ========================
//  辅助函数
// ========================
function buildFullHtml(bodyHtml) {
  return '<!DOCTYPE html>\n' +
    '<html lang="zh-CN">\n' +
    '<head>\n' +
    '  <meta charset="UTF-8">\n' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
    '  <title>导出文档</title>\n' +
    '  <style>\n' +
    '    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; max-width: 860px; margin: 2rem auto; padding: 0 1.5rem; line-height: 1.8; color: #1a1b26; background: #fff; }\n' +
    '    pre { background: #1b1d29; color: #c0caf5; padding: 16px; overflow-x: auto; border-radius: 8px; }\n' +
    '    code { background: #eee; padding: 2px 6px; border-radius: 4px; font-size: 0.9em; }\n' +
    '    pre code { background: none; color: #c0caf5; }\n' +
    '    table { border-collapse: collapse; width: 100%; margin: 1em 0; }\n' +
    '    th, td { border: 1px solid #ccc; padding: 8px 12px; }\n' +
    '    th { background: #f5f5f5; font-weight: 600; }\n' +
    '    blockquote { border-left: 4px solid #7aa2f7; padding-left: 16px; color: #555; margin: 1em 0; font-style: italic; }\n' +
    '    img, video { max-width: 100%; border-radius: 8px; margin: 8px 0; }\n' +
    '    a { color: #7aa2f7; }\n' +
    '    h1, h2, h3, h4 { color: #1a1b26; margin-top: 1.5em; }\n' +
    '    hr { border: none; border-top: 1px solid #ddd; margin: 2em 0; }\n' +
    '    @media (prefers-color-scheme: dark) {\n' +
    '      body { background: #1a1b26; color: #c0caf5; }\n' +
    '      h1, h2, h3, h4 { color: #7aa2f7; }\n' +
    '      blockquote { color: #565f89; }\n' +
    '      table, th, td { border-color: #292e42; }\n' +
    '      th { background: #1e2030; }\n' +
    '    }\n' +
    '  </style>\n' +
    '</head>\n' +
    '<body>' + bodyHtml + '</body>\n' +
    '</html>';
}

function fileToDataURL(file) {
  return new Promise(function(resolve, reject) {
    var r = new FileReader();
    r.onload = function() { resolve(r.result); };
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function downloadBlob(content, filename, mimeType) {
  if (mimeType === undefined) mimeType = 'text/html;charset=utf-8';
  var blob = new Blob([content], { type: mimeType });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ========================
//  Toast & Info
// ========================
function toast(msg, dur) {
  if (dur === undefined) dur = 2000;
  var el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(function() { el.remove(); }, dur);
}

function updateInfo() {
  var mdMode = paneModes.md === 'render' ? '阅读' : '源码';
  var htmlMode = paneModes.html === 'render' ? '阅读' : '源码';
  if (fileInfo) fileInfo.textContent = 'MD[' + mdMode + '] ' + mdSource.value.length + '字符 · HTML[' + htmlMode + '] ' + htmlSource.value.length + '字符';
}

// ========================
//  点击其他地方关闭下拉菜单
// ========================
document.addEventListener('click', function() {
  importMenu.classList.add('hidden');
  exportMenu.classList.add('hidden');
  templateMenu.classList.add('hidden');
});

// ========================
//  Word 导出辅助：HTML → Docx 段落
// ========================
function htmlToDocxParagraphs(html) {
  var parser = new DOMParser();
  var doc = parser.parseFromString(html, 'text/html');
  var els = doc.body.children;
  var paragraphs = [];
  for (var i = 0; i < els.length; i++) {
    var p = nodeToParagraph(els[i]);
    if (p) {
      if (Array.isArray(p)) paragraphs = paragraphs.concat(p);
      else paragraphs.push(p);
    }
  }
  return paragraphs;
}

function nodeToParagraph(node) {
  if (!node) return null;
  var tag = node.nodeName.toLowerCase();
  if (tag === 'h1') return headingPara(node, 1);
  if (tag === 'h2') return headingPara(node, 2);
  if (tag === 'h3') return headingPara(node, 3);
  if (tag === 'h4') return headingPara(node, 4);
  if (tag === 'p')  return para(node);
  if (tag === 'blockquote') return blockquotePara(node);
  if (tag === 'pre') return codePara(node);
  if (tag === 'hr') return new Paragraph({ text: '', border: { bottom: { color: 'auto', space: 1, style: 'single', size: 6 } }, spacing: { before: 200, after: 200 } });
  if (tag === 'ul') return listPara(node, false);
  if (tag === 'ol') return listPara(node, true);
  if (tag === 'table') return tablePara(node);
  return para(node);
}

function headingPara(node, level) {
  var text = cleanText(node);
  var hlMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4 };
  return new Paragraph({ text: text, heading: hlMap[level] || HeadingLevel.HEADING_4, spacing: { before: 240, after: 120 } });
}

function para(node) {
  var children = inlineRuns(node);
  if (!children.length) return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 80, after: 80 } });
  return new Paragraph({ children: children, spacing: { before: 80, after: 80 } });
}

function blockquotePara(node) {
  var children = inlineRuns(node);
  if (!children.length) return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 120, after: 120 } });
  return new Paragraph({ children: children, border: { left: { color: '7aa2f7', space: 4, style: 'single', size: 12 } }, indent: { left: 360 }, spacing: { before: 120, after: 120 } });
}

function codePara(node) {
  var text = node.textContent || '';
  return new Paragraph({ children: [new TextRun({ text: text, font: 'Courier New', size: 18, color: 'c0caf5' })], shading: { type: 'clear', color: '1b1d29', fill: '1b1d29' }, spacing: { before: 120, after: 120 } });
}

function listPara(node, ordered) {
  var items = Array.from(node.children);
  var paragraphs = [];
  items.forEach(function(li) {
    var children = inlineRuns(li);
    paragraphs.push(new Paragraph({ children: children, bullet: { level: 0 }, spacing: { before: 60, after: 60 } }));
  });
  return paragraphs;
}

function tablePara(node) {
  var rows = Array.from(node.querySelectorAll('tr'));
  return rows.map(function(row) {
    var cells = Array.from(row.children).map(function(cell) { return (cell.textContent || '').trim(); });
    return new Paragraph({ text: cells.join(' | '), spacing: { before: 60, after: 60 } });
  });
}

function inlineRuns(node) {
  var runs = [];
  var walk = function(n) {
    for (var i = 0; i < n.childNodes.length; i++) {
      var child = n.childNodes[i];
      if (child.nodeType === Node.TEXT_NODE) {
        var t = child.textContent;
        if (t) runs.push(new TextRun({ text: t }));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        var elTag = child.nodeName.toLowerCase();
        var inner = child.textContent || '';
        if (elTag === 'strong' || elTag === 'b') { runs.push(new TextRun({ text: inner, bold: true })); }
        else if (elTag === 'em' || elTag === 'i') { runs.push(new TextRun({ text: inner, italics: true })); }
        else if (elTag === 'u') { runs.push(new TextRun({ text: inner, underline: {} })); }
        else if (elTag === 's' || elTag === 'strike') { runs.push(new TextRun({ text: inner, strike: true })); }
        else if (elTag === 'code') { runs.push(new TextRun({ text: inner, font: 'Courier New', size: 18 })); }
        else if (elTag === 'a') {
          var href = child.getAttribute('href') || '';
          runs.push(new ExternalHyperlink({ children: [new TextRun({ text: inner, color: '7dcfff', underline: {} })], link: href }));
        }
        else if (elTag === 'br') { runs.push(new TextRun({ text: '\n' })); }
        else { walk(child); }
      }
    }
  };
  walk(node);
  return runs;
}

function cleanText(node) {
  return (node.textContent || '').trim();
}

console.log('🧰 MD ↔ HTML 编辑器 已就绪');