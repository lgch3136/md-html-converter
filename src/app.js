// === MD ↔ HTML 编辑器 v2 ===
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

// 转换按钮
const btnSyncMdToHtml = document.getElementById('btnSyncMdToHtml');
const btnSyncHtmlToMd = document.getElementById('btnSyncHtmlToMd');

// 导入
const importBtn    = document.getElementById('btnImport');
const importMenu   = document.getElementById('importMenu');
const importInput  = document.getElementById('importInput');

// 导出
const exportBtn    = document.getElementById('btnExport');
const exportMenu   = document.getElementById('exportMenu');

// 弹窗
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

// 文件上传
const imageInputMd    = document.getElementById('imageInputMd');
const imageInputHtml  = document.getElementById('imageInputHtml');

// ===== 状态 =====
const paneModes = { md: 'render', html: 'render' };
let pendingTarget = null;
let pendingFmt    = null;

// ========================
//  面板模式切换 (修复版)
//  skipContentSync=true 时跳过内容同步，由调用方自行设置内容
// ========================
document.querySelectorAll('.pane-mode-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    setPaneMode(btn.dataset.pane, btn.dataset.mode);
  });
});

function setPaneMode(pane, mode, skipContentSync = false) {
  const prevMode = paneModes[pane];
  paneModes[pane] = mode;

  // 更新按钮 active 状态
  document.querySelectorAll(`.pane-mode-btn[data-pane="${pane}"]`).forEach(b => {
    b.classList.toggle('active', b.dataset.mode === mode);
  });

  if (pane === 'md') {
    if (mode === 'render') {
      // 从源码切到阅读：把源码内容渲染出来
      if (!skipContentSync && prevMode === 'source') {
        mdRender.innerHTML = mdRenderFromMd(mdSource.value);
      }
      mdSource.classList.add('hidden');
      mdToolbar.classList.remove('hidden');
      mdRender.classList.remove('hidden');
    } else {
      // 从阅读切到源码：把阅读内容转回 MD
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
//  互转按钮 (修复版)
// ========================

// MD → HTML：把 MD 源码转成 HTML，右侧显示渲染预览
btnSyncMdToHtml.addEventListener('click', () => {
  const md = mdSource.value.trim();
  if (!md) { toast('⚠️ 请先在左侧输入 Markdown 内容'); return; }

  const html = marked.parse(md);

  // 1. 先把转换结果写入 htmlSource（后续模式切换会读到它）
  htmlSource.value = html;

  // 2. 确保 MD 侧在阅读模式，显示转换后的 HTML 预览
  setPaneMode('md', 'render', true);
  mdRender.innerHTML = html;

  // 3. 确保 HTML 侧在阅读模式，显示同样的预览效果
  setPaneMode('html', 'render', true);
  htmlRender.innerHTML = html;

  updateInfo();
  toast('MD → HTML ✅ 右侧为预览效果');
});

// HTML → MD
btnSyncHtmlToMd.addEventListener('click', () => {
  const html = htmlSource.value.trim();
  if (!html) { toast('⚠️ 请先在右侧输入 HTML 内容'); return; }

  const md = turndown.turndown(html);

  // 1. 先把转换结果写入 mdSource
  mdSource.value = md;

  // 2. 确保 MD 侧在阅读模式，显示 HTML 的渲染效果
  setPaneMode('md', 'render', true);
  mdRender.innerHTML = html;

  // 3. 确保 HTML 侧在阅读模式
  setPaneMode('html', 'render', true);
  htmlRender.innerHTML = html;

  updateInfo();
  toast('HTML → MD ✅');
});

// ========================
//  富文本格式化
// ========================
function setupToolbar(toolbarEl, renderEl) {
  toolbarEl.querySelectorAll('.fmt-btn').forEach(btn => {
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
      if (sel && sel.toString()) document.execCommand('insertHTML', false, `<code>${sel.toString()}</code>`);
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

document.querySelectorAll('.modal-overlay').forEach(overlay => {
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
    const m = url.match(/BV[\w]+/i);
    src = m ? `https://player.bilibili.com/player.html?bvid=${m[0]}&autoplay=0` : url;
  } else if (url.includes('youtube.com') || url.includes('youtu.be')) {
    const m = url.match(/(?:v=|youtu\.be\/|embed\/)([\w-]{11})/i);
    src = m ? `https://www.youtube.com/embed/${m[1]}` : url;
  }
  const iframeHtml = `<div style="position:relative;width:100%;padding-bottom:56.25%;height:0;overflow:hidden;margin:12px 0;"><iframe src="${src}" style="position:absolute;inset:0;width:100%;height:100%;border:none;border-radius:8px;" allowfullscreen></iframe></div>`;
  pendingTarget.focus();
  document.execCommand('insertHTML', false, iframeHtml);
  closeModal(videoModal);
};

document.getElementById('btnConfirmTable').onclick = () => {
  const rows = parseInt(tableRows.value) || 3;
  const cols = parseInt(tableCols.value) || 3;
  if (!pendingTarget) return;
  if (pendingTarget === mdRender) {
    let md = '\n|';
    for (let c = 0; c < cols; c++) md += ` 列${c+1} |`;
    md += '\n|';
    for (let c = 0; c < cols; c++) md += ' --- |';
    md += '\n';
    for (let r = 0; r < rows - 1; r++) { md += '|'; for (let c = 0; c < cols; c++) md += ' 内容 |'; md += '\n'; }
    document.execCommand('insertHTML', false, marked.parse(md));
  } else {
    let t = '<table style="width:100%;border-collapse:collapse;margin:12px 0;">';
    t += '<thead><tr>';
    for (let c = 0; c < cols; c++) t += `<th style="border:1px solid #292e42;padding:8px 12px;background:#1e2030;color:#7aa2f7;">列${c+1}</th>`;
    t += '</tr></thead><tbody>';
    for (let r = 0; r < rows - 1; r++) { t += '<tr>'; for (let c = 0; c < cols; c++) t += '<td style="border:1px solid #292e42;padding:8px 12px;">内容</td>'; t += '</tr>'; }
    document.execCommand('insertHTML', false, t + '</tbody></table>');
  }
  closeModal(tableModal);
};

document.getElementById('btnConfirmCodeBlock').onclick = () => {
  const lang = codeLang.value.trim();
  const code = codeContent.value;
  if (!pendingTarget) return;
  if (pendingTarget === mdRender) {
    const fence = lang ? `\`\`\`${lang}\n${code}\n\`\`\`` : `\`\`\`\n${code}\n\`\`\``;
    document.execCommand('insertHTML', false, marked.parse(fence));
  } else {
    document.execCommand('insertHTML', false, `<pre style="background:#1b1d29;border:1px solid #292e42;border-radius:8px;padding:12px 16px;overflow-x:auto;margin:12px 0;"><code>${escapeHtml(code)}</code></pre>`);
  }
  closeModal(codeBlockModal);
};

document.getElementById('btnConfirmLink').onclick = () => {
  const text = linkText.value.trim() || '链接';
  const url  = linkUrl.value.trim();
  if (!url || !pendingTarget) return;
  pendingTarget.focus();
  document.execCommand('insertHTML', false, `<a href="${url}" target="_blank" rel="noopener" style="color:#7dcfff;">${text}</a>`);
  closeModal(linkModal);
};

document.getElementById('btnConfirmImage').onclick = async () => {
  const file = imgFileInput.files[0];
  if (!file || !pendingTarget) return;
  const dataUrl = await fileToDataURL(file);
  const alt = file.name.replace(/\.[^.]+$/, '');
  pendingTarget.focus();
  if (pendingTarget === mdRender) {
    const md = `![${alt}](data:image;base64,${dataUrl.split(',')[1]})`;
    document.execCommand('insertHTML', false, marked.parse(md));
  } else {
    document.execCommand('insertHTML', false, `<img src="${dataUrl}" alt="${alt}" style="max-width:100%;border-radius:8px;margin:8px 0;display:block;" />`);
  }
  closeModal(imageModal);
  fileInfo.textContent = `已插入: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  toast(`🖼️ ${file.name}`);
};

// ========================
//  文件上传（工具栏按钮）
// ========================
imageInputMd.addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  await uploadFile(file, 'md');
  imageInputMd.value = '';
});

imageInputHtml.addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  await uploadFile(file, 'html');
  imageInputHtml.value = '';
});

async function uploadFile(file, pane) {
  if (file.name.match(/\.(md|markdown|txt|html)$/i)) {
    const text = await file.text();
    if (pane === 'md') {
      mdSource.value = text;
      if (paneModes.md === 'render') mdRender.innerHTML = mdRenderFromMd(text);
    } else {
      htmlSource.value = text;
      if (paneModes.html === 'render') htmlRender.innerHTML = text;
    }
    fileInfo.textContent = `已加载: ${file.name}`;
    toast(`📄 ${file.name}`);
  } else {
    const dataUrl = await fileToDataURL(file);
    const alt = file.name.replace(/\.[^.]+$/, '');
    const renderEl = pane === 'md' ? mdRender : htmlRender;
    renderEl.focus();
    if (pane === 'md') {
      const md = `![${alt}](data:image;base64,${dataUrl.split(',')[1]})`;
      document.execCommand('insertHTML', false, marked.parse(md));
    } else {
      document.execCommand('insertHTML', false, `<img src="${dataUrl}" alt="${alt}" style="max-width:100%;border-radius:8px;margin:8px 0;display:block;" />`);
    }
    fileInfo.textContent = `已插入: ${file.name}`;
    toast(`🖼️ ${file.name}`);
  }
  updateInfo();
}

// ========================
//  粘贴图片 (Ctrl+V)
// ========================
[mdRender, htmlRender].forEach(el => {
  el.addEventListener('paste', async e => {
    const items = (e.clipboardData || e.originalEvent?.clipboardData)?.items;
    if (!items) return;
    for (const item of items) {
      if (item.type.startsWith('image/')) {
        e.preventDefault();
        const file = item.getAsFile();
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
    const file = e.dataTransfer.files[0];
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
    const pane = getActivePane();
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
// 切换导入菜单
importBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  importMenu.classList.toggle('hidden');
  exportMenu.classList.add('hidden');
});

// 导入菜单项点击 → 触发文件选择
importMenu.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', (e) => {
    e.stopPropagation();
    const format = item.dataset.format;
    // 设置 accept 属性
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

importInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const format = importInput.dataset.importFormat || 'auto';
  await handleImport(file, format);
  importInput.value = '';
});

async function handleImport(file, format) {
  toast('正在导入…');
  try {
    if (format === 'word' || file.name.match(/\.(docx|doc)$/i)) {
      // Word 导入
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      const html = result.value;
      if (result.messages.length) console.warn('mammoth messages:', result.messages);
      const md = turndown.turndown(html);
      mdSource.value = md;
      htmlSource.value = html;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html;
      htmlRender.innerHTML = html;
      fileInfo.textContent = `已导入 Word: ${file.name}`;
      toast(`📥 已导入 Word: ${file.name}`);
    } else if (format === 'md' || file.name.match(/\.(md|markdown)$/i)) {
      // MD 导入
      const text = await file.text();
      const html = marked.parse(text);
      mdSource.value = text;
      htmlSource.value = html;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html;
      htmlRender.innerHTML = html;
      fileInfo.textContent = `已导入 Markdown: ${file.name}`;
      toast(`📥 已导入 MD: ${file.name}`);
    } else if (format === 'html' || file.name.match(/\.(html|htm)$/i)) {
      // HTML 导入
      const html = await file.text();
      const md = turndown.turndown(html);
      mdSource.value = md;
      htmlSource.value = html;
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = html;
      htmlRender.innerHTML = html;
      fileInfo.textContent = `已导入 HTML: ${file.name}`;
      toast(`📥 已导入 HTML: ${file.name}`);
    } else if (format === 'txt' || file.name.match(/\.txt$/i)) {
      // TXT 导入
      const text = await file.text();
      mdSource.value = text;
      htmlSource.value = escapeHtml(text).replace(/\n/g, '<br>');
      setPaneMode('md', 'render', true);
      setPaneMode('html', 'render', true);
      mdRender.innerHTML = `<p>${escapeHtml(text).replace(/\n/g, '<br>')}</p>`;
      htmlRender.innerHTML = `<p>${escapeHtml(text).replace(/\n/g, '<br>')}</p>`;
      fileInfo.textContent = `已导入文本: ${file.name}`;
      toast(`📥 已导入 TXT: ${file.name}`);
    } else {
      // 未知格式，当作 MD 尝试
      try {
        const text = await file.text();
        const html = marked.parse(text);
        mdSource.value = text;
        htmlSource.value = html;
        setPaneMode('md', 'render', true);
        setPaneMode('html', 'render', true);
        mdRender.innerHTML = html;
        htmlRender.innerHTML = html;
        fileInfo.textContent = `已导入: ${file.name}`;
        toast(`📥 已导入: ${file.name}`);
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
exportBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  exportMenu.classList.toggle('hidden');
  importMenu.classList.add('hidden');
});

exportMenu.querySelectorAll('.menu-item').forEach(item => {
  item.addEventListener('click', async (e) => {
    e.stopPropagation();
    exportMenu.classList.add('hidden');
    const format = item.dataset.format;
    await handleExport(format);
  });
});

async function handleExport(format) {
  const md = mdSource.value.trim();
  if (!md) { toast('⚠️ 没有内容可导出'); return; }

  try {
    const htmlBody = marked.parse(md);

    switch (format) {
      case 'pdf':
        await exportPdf(htmlBody);
        break;
      case 'word':
        await exportWord(htmlBody);
        break;
      case 'html':
        exportStandaloneHtml(htmlBody);
        break;
      case 'md':
        exportMarkdown(md);
        break;
      case 'txt':
        exportText(htmlBody);
        break;
      default:
        toast('未知导出格式');
    }
  } catch (err) {
    console.error(err);
    toast(`导出 ${format.toUpperCase()} 失败 ❌`);
  }
}

// --- PDF 导出 ---
async function exportPdf(htmlBody) {
  toast('正在生成 PDF…');
  // 动态加载 html2pdf
  const html2pdf = await loadHtml2Pdf();
  const fullHtml = buildFullHtml(htmlBody);

  // 创建临时 iframe 来渲染
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.left = '-9999px';
  iframe.style.width = '800px';
  iframe.style.height = '600px';
  document.body.appendChild(iframe);

  iframe.contentDocument.open();
  iframe.contentDocument.write(fullHtml);
  iframe.contentDocument.close();

  // 等待图片加载
  await new Promise(resolve => setTimeout(resolve, 500));

  const opt = {
    margin: [10, 10, 10, 10],
    filename: 'document.pdf',
    image: { type: 'jpeg', quality: 0.95 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  await html2pdf().set(opt).from(iframe.contentDocument.body).save();
  document.body.removeChild(iframe);
  toast('PDF 已导出 📄');
}

// 动态加载 html2pdf.js
let html2pdfModule = null;
async function loadHtml2Pdf() {
  if (html2pdfModule) return html2pdfModule;
  // html2pdf 挂载到 window.html2pdf
  if (window.html2pdf) return window.html2pdf;

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.3/html2pdf.bundle.min.js';
    script.onload = () => {
      html2pdfModule = window.html2pdf;
      resolve(window.html2pdf);
    };
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// --- Word 导出 ---
async function exportWord(htmlBody) {
  toast('正在生成 Word…');
  const paragraphs = htmlToDocxParagraphs(htmlBody);
  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });
  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, 'document.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  toast('Word 已导出 📤');
}

// --- 独立 HTML 导出 ---
function exportStandaloneHtml(htmlBody) {
  const full = buildFullHtml(htmlBody);
  downloadBlob(full, 'document.html', 'text/html;charset=utf-8');
  toast('HTML 已导出 🌐');
}

// --- MD 导出 ---
function exportMarkdown(md) {
  downloadBlob(md, 'document.md', 'text/markdown;charset=utf-8');
  toast('Markdown 已导出 📝');
}

// --- TXT 导出（纯文本） ---
function exportText(htmlBody) {
  const div = document.createElement('div');
  div.innerHTML = htmlBody;
  const text = div.textContent || '';
  downloadBlob(text, 'document.txt', 'text/plain;charset=utf-8');
  toast('纯文本已导出 📃');
}

// ========================
//  辅助函数
// ========================
function buildFullHtml(bodyHtml) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>导出文档</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      max-width: 860px;
      margin: 2rem auto;
      padding: 0 1.5rem;
      line-height: 1.8;
      color: #1a1b26;
      background: #fff;
    }
    pre { background: #1b1d29; color: #c0caf5; padding: 16px; overflow-x: auto; border-radius: 8px; }
    code { background: #eee; padding: 2px 6px; border-radius: 4px; font-size: 0.9em; }
    pre code { background: none; color: #c0caf5; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    th, td { border: 1px solid #ccc; padding: 8px 12px; }
    th { background: #f5f5f5; font-weight: 600; }
    blockquote { border-left: 4px solid #7aa2f7; padding-left: 16px; color: #555; margin: 1em 0; font-style: italic; }
    img, video { max-width: 100%; border-radius: 8px; margin: 8px 0; }
    a { color: #7aa2f7; }
    h1, h2, h3, h4 { color: #1a1b26; margin-top: 1.5em; }
    hr { border: none; border-top: 1px solid #ddd; margin: 2em 0; }
    @media (prefers-color-scheme: dark) {
      body { background: #1a1b26; color: #c0caf5; }
      h1, h2, h3, h4 { color: #7aa2f7; }
      blockquote { color: #565f89; }
      table, th, td { border-color: #292e42; }
      th { background: #1e2030; }
    }
  </style>
</head>
<body>${bodyHtml}</body>
</html>`;
}

function fileToDataURL(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function downloadBlob(content, filename, mimeType = 'text/html;charset=utf-8') {
  const blob = new Blob([content], { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ========================
//  Toast & Info
// ========================
function toast(msg, dur = 2000) {
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), dur);
}

function updateInfo() {
  const mdMode  = paneModes.md  === 'render' ? '阅读' : '源码';
  const htmlMode = paneModes.html === 'render' ? '阅读' : '源码';
  fileInfo.textContent = `MD[${mdMode}] ${mdSource.value.length}字符 · HTML[${htmlMode}] ${htmlSource.value.length}字符`;
}

// ========================
//  点击其他地方关闭下拉菜单
// ========================
document.addEventListener('click', () => {
  importMenu.classList.add('hidden');
  exportMenu.classList.add('hidden');
});

// ========================
//  Word 导出辅助：HTML → Docx 段落
// ========================
function htmlToDocxParagraphs(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const els = doc.body.children;
  const paragraphs = [];
  for (const el of els) {
    const p = nodeToParagraph(el);
    if (p) {
      if (Array.isArray(p)) paragraphs.push(...p);
      else paragraphs.push(p);
    }
  }
  return paragraphs;
}

function nodeToParagraph(node) {
  if (!node) return null;
  switch (node.nodeName.toLowerCase()) {
    case 'h1': return headingPara(node, 1);
    case 'h2': return headingPara(node, 2);
    case 'h3': return headingPara(node, 3);
    case 'h4': return headingPara(node, 4);
    case 'p':  return para(node);
    case 'blockquote': return blockquotePara(node);
    case 'pre': return codePara(node);
    case 'hr': return new Paragraph({ text: '', border: { bottom: { color: 'auto', space: 1, style: 'single', size: 6 } }, spacing: { before: 200, after: 200 } });
    case 'ul': return listPara(node, false);
    case 'ol': return listPara(node, true);
    case 'table': return tablePara(node);
    default: return para(node);
  }
}

function headingPara(node, level) {
  const text = cleanText(node);
  const hlMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4 };
  return new Paragraph({ text, heading: hlMap[level] || HeadingLevel.HEADING_4, spacing: { before: 240, after: 120 } });
}

function para(node) {
  const children = inlineRuns(node);
  if (!children.length) return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 80, after: 80 } });
  return new Paragraph({ children, spacing: { before: 80, after: 80 } });
}

function blockquotePara(node) {
  const children = inlineRuns(node);
  if (!children.length) return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 120, after: 120 } });
  return new Paragraph({ children, border: { left: { color: '7aa2f7', space: 4, style: 'single', size: 12 } }, indent: { left: 360 }, spacing: { before: 120, after: 120 } });
}

function codePara(node) {
  const text = node.textContent || '';
  return new Paragraph({ children: [new TextRun({ text, font: 'Courier New', size: 18, color: 'c0caf5' })], shading: { type: 'clear', color: '1b1d29', fill: '1b1d29' }, spacing: { before: 120, after: 120 } });
}

function listPara(node, ordered) {
  const items = Array.from(node.children);
  const paragraphs = [];
  items.forEach((li, i) => {
    const children = inlineRuns(li);
    paragraphs.push(new Paragraph({ children, bullet: { level: 0 }, spacing: { before: 60, after: 60 } }));
  });
  return paragraphs;
}

function tablePara(node) {
  const rows = Array.from(node.querySelectorAll('tr'));
  return rows.map(row => {
    const cells = Array.from(row.children).map(cell => (cell.textContent || '').trim());
    return new Paragraph({ text: cells.join(' | '), spacing: { before: 60, after: 60 } });
  });
}

function inlineRuns(node) {
  const runs = [];
  const walk = (n) => {
    for (const child of n.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent;
        if (text) runs.push(new TextRun({ text }));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const tag = child.nodeName.toLowerCase();
        const inner = child.textContent || '';
        switch (tag) {
          case 'strong': case 'b': runs.push(new TextRun({ text: inner, bold: true })); break;
          case 'em': case 'i':     runs.push(new TextRun({ text: inner, italics: true })); break;
          case 'u':                runs.push(new TextRun({ text: inner, underline: {} })); break;
          case 's': case 'strike': runs.push(new TextRun({ text: inner, strike: true })); break;
          case 'code':             runs.push(new TextRun({ text: inner, font: 'Courier New', size: 18 })); break;
          case 'a': {
            const href = child.getAttribute('href') || '';
            runs.push(new ExternalHyperlink({ children: [new TextRun({ text: inner, color: '7dcfff', underline: {} })], link: href }));
            break;
          }
          case 'br': runs.push(new TextRun({ text: '\n' })); break;
          default: walk(child);
        }
      }
    }
  };
  walk(node);
  return runs;
}

function cleanText(node) {
  return (node.textContent || '').trim();
}

// ========================
//  初始化
// ========================
const defaultMd = `# 欢迎使用 MD ↔ HTML 编辑器 v2 ✨

支持**富媒体**的双栏文档编辑器，左右均可独立切换「阅读 / 源码」模式。

## 功能概览

| 功能 | 状态 |
|------|------|
| 两侧独立阅读/源码切换 | ✅ |
| 富文本编辑（加粗/斜体/标题等）| ✅ |
| 粘贴/拖入图片 & 视频嵌入 | ✅ |
| MD ↔ HTML 双向互转 | ✅ |
| 导入 MD / HTML / TXT / Word | ✅ |
| 导出 PDF / Word / HTML / MD / TXT | ✅ |
| Vercel 一键部署 | ✅ |

## 快速开始

> 在左侧输入 Markdown，点击中间「转 HTML」按钮，
> 右侧即可看到渲染预览效果！

\`\`\`javascript
console.log("Hello, MD ↔ HTML!");
\`\`\`

试着用工具栏的格式化按钮，或 **Ctrl+B** / **Ctrl+I** 快捷键编辑格式 😊
`;

mdSource.value = defaultMd;
const defaultHtml = marked.parse(defaultMd);
htmlSource.value = defaultHtml;
mdRender.innerHTML = defaultHtml;
htmlRender.innerHTML = defaultHtml;
updateInfo();

console.log('🧰 MD ↔ HTML 编辑器 v2 已就绪');
