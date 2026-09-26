import {
  createIcons,
  Save,
  Undo2,
  Redo2,
  Printer,
  Download,
  FolderOpen,
  Table,
  Image as ImageIcon,
  Scissors,
  Minus,
  Link as LinkIcon,
  Maximize,
  Minimize,
  Move,
  File,
  FileMinus,
  FileText,
  FileSearch,
  Edit3,
  ArrowLeftRight,
  Info,
  Scan,
  Maximize2,
  Expand,
  ArrowLeft,
  UploadCloud
} from 'lucide';

import { initWordEditor, measureLineHeightFactor } from './editor.js';
import { DocxService } from './docx-service.js';

// Application State
const state = {
  editor: null,
  currentFileName: 'Tài liệu 1',
  lastRawDocxBuffer: null,
  isDocxPreviewMode: false,
  autoSaveTimer: null,
  zoomLevel: 100,
  currentPageLayout: null
};

// DOM Elements
const docTitleInput = document.getElementById('docTitleInput');
const saveStatusIndicator = document.getElementById('saveStatusIndicator');
const docxFileInput = document.getElementById('docxFileInput');
const imageFileInput = document.getElementById('imageFileInput');
const wordPage = document.getElementById('wordPage');
const editorViewport = document.getElementById('editorViewport');
const previewViewport = document.getElementById('previewViewport');
const docxRenderTarget = document.getElementById('docxRenderTarget');
const previewFileName = document.getElementById('previewFileName');
const rulerContainer = document.getElementById('rulerContainer');
const rulerBody = document.getElementById('rulerBody');
const backstageModal = document.getElementById('backstageModal');
const wordToast = document.getElementById('wordToast');
const toastMessage = document.getElementById('toastMessage');

// Status Bar Elements
const sbPageCount = document.getElementById('sbPageCount');
const sbWordCount = document.getElementById('sbWordCount');
const sbCharCount = document.getElementById('sbCharCount');
const sbMessage = document.getElementById('sbMessage');
const zoomSlider = document.getElementById('zoomSlider');
const zoomValueText = document.getElementById('zoomValueText');

/**
 * Initialize Lucide Icons
 */
function initIcons() {
  createIcons({
    icons: {
      Save,
      Undo2,
      Redo2,
      Printer,
      Download,
      FolderOpen,
      Table,
      Image: ImageIcon,
      Scissors,
      Minus,
      Link: LinkIcon,
      Maximize,
      Minimize,
      Move,
      File,
      FileMinus,
      FileText,
      FileSearch,
      Edit3,
      ArrowLeftRight,
      Info,
      Scan,
      Maximize2,
      Expand,
      ArrowLeft,
      UploadCloud
    }
  });
}

/**
 * Show Toast Notification
 */
function showToast(message, isError = false) {
  toastMessage.textContent = message;
  wordToast.classList.toggle('error', isError);
  wordToast.classList.add('show');
  setTimeout(() => {
    wordToast.classList.remove('show');
  }, 3200);
}

/**
 * Set Save Status Indicator
 */
function setSaveStatus(status) {
  saveStatusIndicator.className = `save-status-badge ${status}`;
  const textElem = saveStatusIndicator.querySelector('.status-text');
  if (status === 'saved') {
    textElem.textContent = 'Đã lưu';
  } else if (status === 'saving') {
    textElem.textContent = 'Đang lưu...';
  } else {
    textElem.textContent = 'Chưa lưu';
  }
}

/**
 * Apply page layout (size + margins) from detected DOCX pgSz / pgMar values
 * @param {{ pageWidthPx, pageHeightPx, marginTopPx, marginRightPx, marginBottomPx, marginLeftPx, orientation }} layout
 */
function applyPageLayout(layout) {
  if (!layout) return;

  const { pageWidthPx, pageHeightPx, marginTopPx, marginRightPx, marginBottomPx, marginLeftPx, orientation } = layout;

  // Apply explicit pixel dimensions to override fixed CSS classes
  wordPage.style.width = `${pageWidthPx}px`;
  wordPage.style.minHeight = `${pageHeightPx}px`;
  wordPage.style.paddingTop = `${marginTopPx}px`;
  wordPage.style.paddingRight = `${marginRightPx}px`;
  wordPage.style.paddingBottom = `${marginBottomPx}px`;
  wordPage.style.paddingLeft = `${marginLeftPx}px`;
  wordPage.style.setProperty('--page-margin-top', `${marginTopPx}px`);
  wordPage.style.setProperty('--page-margin-right', `${marginRightPx}px`);
  wordPage.style.setProperty('--page-margin-bottom', `${marginBottomPx}px`);
  wordPage.style.setProperty('--page-margin-left', `${marginLeftPx}px`);

  // Remove static margin classes - layout is now fully dynamic
  wordPage.classList.remove('margin-normal', 'margin-narrow', 'margin-wide');

  // Apply orientation class
  wordPage.classList.remove('portrait', 'landscape');
  if (orientation) wordPage.classList.add(orientation);

  // Re-render ruler with new page width and margins
  renderRulerMarks(pageWidthPx, marginLeftPx, marginRightPx);

  console.info(`[PageLayout] ${pageWidthPx}x${pageHeightPx}px | margins T:${marginTopPx} R:${marginRightPx} B:${marginBottomPx} L:${marginLeftPx}px`);
}

/**
 * Draw Horizontal Ruler Markings (Word cm / mm style)
 */
function renderRulerMarks(rulerWidthPx = 794, marginLeftPx = 96, marginRightPx = 96) {
  rulerBody.innerHTML = '';
  // A4 printable width in cm ~ 16cm between standard 2.54cm margins
  // 1 cm ~ 37.8px at 96 DPI
  const pxPerCm = 37.795;
  const totalCm = Math.floor(rulerWidthPx / pxPerCm);

  for (let cm = 0; cm <= totalCm; cm++) {
    const leftPx = cm * pxPerCm;
    if (leftPx > rulerWidthPx) break;

    const cmMark = document.createElement('div');
    cmMark.className = 'ruler-cm-mark';
    cmMark.style.left = `${leftPx}px`;
    cmMark.textContent = cm.toString();
    rulerBody.appendChild(cmMark);

    // mm marks
    for (let mm = 1; mm < 10; mm++) {
      const mmPx = leftPx + (mm * (pxPerCm / 10));
      if (mmPx >= rulerWidthPx) break;
      const mmMark = document.createElement('div');
      mmMark.className = 'ruler-mm-mark';
      mmMark.style.left = `${mmPx}px`;
      if (mm === 5) {
        mmMark.style.height = '6px';
      }
      rulerBody.appendChild(mmMark);
    }
  }

  // Margin Indicators
  const leftIndicator = document.createElement('div');
  leftIndicator.className = 'ruler-margin-indicator';
  leftIndicator.style.left = `${marginLeftPx}px`;
  rulerBody.appendChild(leftIndicator);

  const rightIndicator = document.createElement('div');
  rightIndicator.className = 'ruler-margin-indicator';
  rightIndicator.style.right = `${marginRightPx}px`;
  rulerBody.appendChild(rightIndicator);
}

/**
 * Zoom Handling
 */
function setZoom(percent) {
  percent = Math.max(50, Math.min(180, percent));
  state.zoomLevel = percent;
  zoomSlider.value = percent;
  zoomValueText.textContent = `${percent}%`;
  editorViewport.style.transform = `scale(${percent / 100})`;
  
  if (state.isDocxPreviewMode) {
    previewViewport.style.transform = `scale(${percent / 100})`;
    previewViewport.style.transformOrigin = 'top center';
  }
}

/**
 * Switch View Mode: Editor vs DOCX Preview
 */
function switchViewMode(mode) {
  if (mode === 'preview') {
    state.isDocxPreviewMode = true;
    editorViewport.style.display = 'none';
    previewViewport.style.display = 'flex';
    document.querySelectorAll('.ribbon-tab').forEach(t => t.classList.remove('active'));
    document.querySelector('.preview-tab-btn').classList.add('active');
    document.querySelectorAll('.ribbon-panel').forEach(p => p.classList.remove('active'));
    document.getElementById('panelPreview').classList.add('active');
    sbMessage.textContent = 'Đang xem trước nguyên bản DOCX (docx-preview)';
  } else {
    state.isDocxPreviewMode = false;
    previewViewport.style.display = 'none';
    editorViewport.style.display = 'flex';
    document.querySelectorAll('.ribbon-tab').forEach(t => t.classList.remove('active'));
    document.querySelector('[data-target="panelHome"]').classList.add('active');
    document.querySelectorAll('.ribbon-panel').forEach(p => p.classList.remove('active'));
    document.getElementById('panelHome').classList.add('active');
    sbMessage.textContent = 'Sẵn sàng soạn thảo';
  }
}

/**
 * Load Document into CKEditor
 */
function loadHtmlIntoEditor(htmlContent, title, font, fontSize, lineSpacing) {
  if (state.editor) {
    // Match Word's line metrics for the document font before content is rendered
    const editorRoot = document.getElementById('editor');
    if (editorRoot && font) {
      const family = `'${font}', 'Times New Roman', Calibri, serif`;
      editorRoot.style.setProperty('--word-lh-factor', String(measureLineHeightFactor(family)));
    }
    if (editorRoot && fontSize) {
      editorRoot.style.fontSize = `${fontSize}pt`;
    }

    state.editor.setData(htmlContent);
    if (title) {
      state.currentFileName = title;
      docTitleInput.value = title;
      document.getElementById('bsDocName').textContent = `${title}.docx`;
    }

    // Apply document font and font-size to editor and update ribbon
    if (font) {
      state.currentFont = font;
      const editorEl = document.getElementById('editor');
      if (editorEl) {
        editorEl.style.fontFamily = `'${font}', 'Times New Roman', Calibri, serif`;
      }
      setTimeout(() => {
        const familyDropdown = document.querySelector('.ck-font-family-dropdown');
        if (familyDropdown) {
          const label = familyDropdown.querySelector('.ck-button__label');
          if (label) label.textContent = font;
        }
      }, 100);
    }

    if (fontSize) {
      state.currentFontSize = fontSize;
      setTimeout(() => {
        const sizeDropdown = document.querySelector('.ck-font-size-dropdown');
        if (sizeDropdown) {
          const label = sizeDropdown.querySelector('.ck-button__label');
          if (label) label.textContent = String(fontSize);
        }
      }, 100);
    }

    if (lineSpacing) {
      state.currentLineSpacing = lineSpacing;
      const currentLineSpacingText = document.getElementById('currentLineSpacingText');
      if (currentLineSpacingText) {
        currentLineSpacingText.textContent = lineSpacing;
      }
      document.querySelectorAll('[data-lineheight]').forEach(i => {
        i.classList.toggle('active', i.dataset.lineheight === String(lineSpacing));
      });
    }

    switchViewMode('editor');
    const fontInfo = font ? ` với phông "${font}" (${fontSize || 14}pt, giãn dòng ${lineSpacing || '1.15'})` : '';
    showToast(`Đã nạp văn bản "${title || 'Tài liệu'}"${fontInfo}!`);
  }
}

/**
 * Save Document (Export to DOCX & Auto-save)
 */
async function saveDocument() {
  if (!state.editor) return;
  setSaveStatus('saving');
  sbMessage.textContent = 'Đang xuất tài liệu Word (.docx)...';

  const htmlData = state.editor.getData();
  const rawTitle = docTitleInput.value.trim() || state.currentFileName || 'TaiLieu';
  const cleanTitle = rawTitle.replace(/\.docx$/i, '').trim() || 'TaiLieu';
  
  try {
    // Build margin options in TWIPs (1px at 96DPI = 15 TWIPs)
    const PX_TO_TWIPS = 1440 / 96;
    const layout = state.currentPageLayout;
    const exportMargins = layout ? {
      top: Math.round(layout.marginTopPx * PX_TO_TWIPS),
      right: Math.round(layout.marginRightPx * PX_TO_TWIPS),
      bottom: Math.round(layout.marginBottomPx * PX_TO_TWIPS),
      left: Math.round(layout.marginLeftPx * PX_TO_TWIPS)
    } : { top: 1440, right: 1440, bottom: 1440, left: 1440 };

    const result = await DocxService.exportToDocx(htmlData, `${cleanTitle}.docx`, {
      orientation: layout?.orientation || 'portrait',
      margins: exportMargins
    });
    if (result && result.cancelled) {
      setSaveStatus('unsaved');
      sbMessage.textContent = 'Đã hủy lưu tệp.';
      return;
    }
    DocxService.saveToLocalStorage(htmlData, cleanTitle);
    setSaveStatus('saved');
    const finalDocName = result?.fileName || `${cleanTitle}.docx`;
    sbMessage.textContent = `Tài liệu "${finalDocName}" đã được lưu thành công!`;
    showToast(`Đã xuất file "${finalDocName}" về máy tính thành công!`);
  } catch (err) {
    console.error('Save error:', err);
    setSaveStatus('unsaved');
    sbMessage.textContent = 'Lỗi khi lưu tài liệu.';
    showToast(`Lỗi khi xuất file: ${err.message}`, true);
  }
}

/**
 * Handle File Selection (Import)
 */
async function handleFileImport(file, targetMode = 'auto') {
  if (!file) return;

  const fileName = file.name;
  sbMessage.textContent = `Đang đọc tệp "${fileName}"...`;

  try {
    const imported = await DocxService.importDocFile(file);
    state.lastRawDocxBuffer = imported.rawBuffer;
    previewFileName.textContent = fileName;
    document.getElementById('btnImportDocxToEditor').disabled = false;

    if (targetMode === 'preview') {
      // Direct docx-preview
      switchViewMode('preview');
      await DocxService.previewDocx(imported.rawBuffer, docxRenderTarget);
      showToast(`Đã mở bản xem trước "${fileName}" chuẩn Word!`);
    } else {
      // Apply detected page dimensions BEFORE loading into editor
      if (imported.pageLayout) {
        state.currentPageLayout = imported.pageLayout;
        applyPageLayout(imported.pageLayout);
      }
      // Load directly into editor preserving detected font, size & line-spacing
      loadHtmlIntoEditor(imported.html, imported.title, imported.font, imported.fontSize, imported.lineSpacing);
    }
  } catch (err) {
    console.error('Import error:', err);
    showToast(err.message, true);
    sbMessage.textContent = 'Lỗi nhập tệp.';
  }
}

/**
 * Setup Main Event Listeners
 */
function setupEventListeners() {
  // Quick Access buttons
  document.getElementById('qaSaveBtn').addEventListener('click', saveDocument);
  document.getElementById('btnQuickExport').addEventListener('click', saveDocument);
  document.getElementById('btnExportDocx').addEventListener('click', () => {
    saveDocument();
    backstageModal.style.display = 'none';
  });

  document.getElementById('qaUndoBtn').addEventListener('click', () => {
    if (state.editor) state.editor.execute('undo');
  });

  document.getElementById('qaRedoBtn').addEventListener('click', () => {
    if (state.editor) state.editor.execute('redo');
  });

  document.getElementById('qaPrintBtn').addEventListener('click', () => window.print());
  document.getElementById('btnExecutePrint').addEventListener('click', () => {
    backstageModal.style.display = 'none';
    setTimeout(() => window.print(), 100);
  });

  // Quick Import buttons
  document.getElementById('btnQuickImport').addEventListener('click', () => {
    docxFileInput.click();
  });

  document.getElementById('btnBrowseFile').addEventListener('click', () => {
    docxFileInput.click();
  });

  document.getElementById('btnOpenDocxForPreview').addEventListener('click', () => {
    docxFileInput.dataset.mode = 'preview';
    docxFileInput.click();
  });

  docxFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      const mode = docxFileInput.dataset.mode || 'auto';
      handleFileImport(file, mode);
      docxFileInput.dataset.mode = '';
      docxFileInput.value = '';
      backstageModal.style.display = 'none';
    }
  });

  // Transfer from Preview to Editor
  document.getElementById('btnBannerTransferToEditor').addEventListener('click', async () => {
    if (state.lastRawDocxBuffer) {
      const blob = new Blob([state.lastRawDocxBuffer]);
      const file = new File([blob], `${previewFileName.textContent || 'TaiLieu'}.docx`);
      await handleFileImport(file, 'editor');
    }
  });

  document.getElementById('btnImportDocxToEditor').addEventListener('click', async () => {
    if (state.lastRawDocxBuffer) {
      const blob = new Blob([state.lastRawDocxBuffer]);
      const file = new File([blob], `${previewFileName.textContent || 'TaiLieu'}.docx`);
      await handleFileImport(file, 'editor');
    }
  });

  document.getElementById('btnSwitchToEditor').addEventListener('click', () => switchViewMode('editor'));
  document.getElementById('btnClosePreview').addEventListener('click', () => switchViewMode('editor'));

  const btnLoadUserDoc = document.getElementById('btnLoadUserDoc');
  if (btnLoadUserDoc) {
    btnLoadUserDoc.addEventListener('click', async () => {
      try {
        const resp = await fetch('/tes1.docx');
        const blob = await resp.blob();
        const file = new File([blob], 'tes1.docx');
        await handleFileImport(file, 'editor');
      } catch (e) {
        showToast('Lỗi nạp tệp: ' + e.message, true);
      }
    });
  }

  document.getElementById('btnBannerDownloadDocx').addEventListener('click', () => {
    if (state.lastRawDocxBuffer) {
      const blob = new Blob([state.lastRawDocxBuffer]);
      DocxService.exportToDocx(state.editor ? state.editor.getData() : '<p></p>', `${state.currentFileName}.docx`);
    }
  });

  // Ribbon Tab Switching
  const ribbonTabs = document.querySelectorAll('.ribbon-tab:not(.file-tab)');
  ribbonTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      ribbonTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      const targetPanelId = tab.dataset.target;
      document.querySelectorAll('.ribbon-panel').forEach(panel => {
        panel.classList.remove('active');
      });

      const activePanel = document.getElementById(targetPanelId);
      if (activePanel) activePanel.classList.add('active');

      if (targetPanelId === 'panelPreview') {
        if (!state.isDocxPreviewMode && state.lastRawDocxBuffer) {
          switchViewMode('preview');
        }
      } else {
        if (state.isDocxPreviewMode) {
          switchViewMode('editor');
        }
      }
    });
  });

  // File Tab & Backstage Modal
  document.getElementById('tabFile').addEventListener('click', () => {
    backstageModal.style.display = 'flex';
  });

  document.getElementById('btnBackstageBack').addEventListener('click', () => {
    backstageModal.style.display = 'none';
  });

  document.getElementById('backstageOverlay').addEventListener('click', () => {
    backstageModal.style.display = 'none';
  });

  // Backstage Navigation
  const bsNavItems = document.querySelectorAll('.backstage-nav-item');
  bsNavItems.forEach(item => {
    item.addEventListener('click', () => {
      bsNavItems.forEach(i => i.classList.remove('active'));
      item.classList.add('active');

      const contentId = item.dataset.content;
      document.querySelectorAll('.backstage-tab-content').forEach(c => c.classList.remove('active'));
      const activeContent = document.getElementById(contentId);
      if (activeContent) activeContent.classList.add('active');

      if (contentId === 'bsSave') {
        saveDocument();
        backstageModal.style.display = 'none';
      }
    });
  });

  // Templates
  document.getElementById('btnNewBlankDoc').addEventListener('click', () => {
    loadHtmlIntoEditor('<p></p>', 'Tài liệu mới');
    backstageModal.style.display = 'none';
  });

  document.getElementById('btnNewReportDoc').addEventListener('click', () => {
    const reportHtml = `
      <h1 style="color: #1f4e79; text-align: center;">BÁO CÁO KẾT QUẢ CÔNG VIỆC</h1>
      <p style="text-align: center; color: #595959;"><em>Ngày lập: ${new Date().toLocaleDateString('vi-VN')}</em></p>
      <hr>
      <h2>1. TỔNG QUAN</h2>
      <p>Báo cáo tóm tắt các hoạt động, tiến độ triển khai dự án trong giai đoạn vừa qua.</p>
      <h2>2. KẾT QUẢ ĐẠT ĐƯỢC</h2>
      <ul>
        <li>Hoàn thành mục tiêu đề ra đúng thời hạn.</li>
        <li>Tối ưu hóa quy trình phối hợp nhóm và tài liệu.</li>
      </ul>
      <h2>3. KẾ HOẠCH TIẾP THEO</h2>
      <p>Tiếp tục duy trì và nâng cấp các tính năng trọng yếu.</p>
    `;
    loadHtmlIntoEditor(reportHtml, 'Báo cáo công việc');
    backstageModal.style.display = 'none';
  });

  // Export HTML
  document.getElementById('btnExportHtml').addEventListener('click', () => {
    if (!state.editor) return;
    const htmlData = state.editor.getData();
    const blob = new Blob([htmlData], { type: 'text/html;charset=utf-8' });
    const fileName = `${docTitleInput.value.trim() || 'TaiLieu'}.html`;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    a.click();
    showToast(`Đã xuất file "${fileName}" thành công!`);
    backstageModal.style.display = 'none';
  });

  // Insert Tab Actions
  document.getElementById('btnInsertTable').addEventListener('click', () => {
    if (state.editor) {
      state.editor.execute('insertTable', { rows: 3, columns: 3 });
      showToast('Đã chèn bảng 3x3 vào tài liệu!');
    }
  });

  document.getElementById('btnInsertImage').addEventListener('click', () => {
    imageFileInput.click();
  });

  imageFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file && state.editor) {
      const reader = new FileReader();
      reader.onload = () => {
        state.editor.execute('insertImage', { source: reader.result });
        showToast('Đã chèn hình ảnh thành công!');
      };
      reader.readAsDataURL(file);
      imageFileInput.value = '';
    }
  });

  document.getElementById('btnInsertPageBreak').addEventListener('click', () => {
    if (state.editor) {
      state.editor.execute('pageBreak');
      showToast('Đã chèn dấu ngắt trang A4!');
    }
  });

  document.getElementById('btnInsertHr').addEventListener('click', () => {
    if (state.editor) {
      state.editor.execute('horizontalLine');
    }
  });

  document.getElementById('btnInsertLink').addEventListener('click', () => {
    if (state.editor) {
      const url = prompt('Nhập địa chỉ URL liên kết:', 'https://');
      if (url) {
        state.editor.execute('link', url);
      }
    }
  });

  // Line Spacing Dropdown (Căn dòng / Giãn dòng)
  const btnLineSpacing = document.getElementById('btnLineSpacing');
  const menuLineSpacing = document.getElementById('menuLineSpacing');
  const currentLineSpacingText = document.getElementById('currentLineSpacingText');

  if (btnLineSpacing && menuLineSpacing) {
    btnLineSpacing.addEventListener('click', (e) => {
      e.stopPropagation();
      menuLineSpacing.classList.toggle('show');
    });

    document.querySelectorAll('[data-lineheight]').forEach(item => {
      item.addEventListener('click', (e) => {
        e.stopPropagation();
        const lh = item.dataset.lineheight;
        currentLineSpacingText.textContent = lh;

        document.querySelectorAll('[data-lineheight]').forEach(i => i.classList.remove('active'));
        item.classList.add('active');

        // Apply line-height directly into CKEditor model blocks
        if (state.editor) {
          state.editor.model.change(writer => {
            const blocks = Array.from(state.editor.model.document.selection.getSelectedBlocks());
            if (blocks.length > 0) {
              for (const block of blocks) {
                writer.setAttribute('lineHeight', lh, block);
              }
            } else {
              const root = state.editor.model.document.getRoot();
              for (const child of root.getChildren()) {
                if (child.is('element')) {
                  writer.setAttribute('lineHeight', lh, child);
                }
              }
            }
          });
        }

        const editorContent = document.getElementById('editor');
        if (editorContent) {
          editorContent.style.lineHeight = `calc(${lh} * var(--word-lh-factor, 1))`;
        }

        menuLineSpacing.classList.remove('show');
        showToast(`Đã đổi giãn cách dòng: ${lh}x`);
      });
    });

    window.addEventListener('click', () => {
      menuLineSpacing.classList.remove('show');
    });
  }

  // Layout Tab Actions: Margins
  document.querySelectorAll('[data-margin]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-margin]').forEach(b => b.classList.remove('active-margin'));
      btn.classList.add('active-margin');
      const margin = btn.dataset.margin;
      wordPage.classList.remove('margin-normal', 'margin-narrow', 'margin-wide');
      wordPage.classList.add(`margin-${margin}`);
      showToast(`Đã đổi lề trang: ${btn.textContent.trim()}`);
    });
  });

  // Layout Tab Actions: Orientation
  document.querySelectorAll('[data-orient]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-orient]').forEach(b => b.classList.remove('active-orient'));
      btn.classList.add('active-orient');
      const orient = btn.dataset.orient;
      wordPage.classList.remove('portrait', 'landscape');
      wordPage.classList.add(orient);
      showToast(`Đã đổi hướng giấy: ${btn.textContent.trim()}`);
    });
  });

  // Layout Tab Actions: Size
  document.querySelectorAll('[data-size]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-size]').forEach(b => b.classList.remove('active-size'));
      btn.classList.add('active-size');
      const size = btn.dataset.size;
      wordPage.classList.remove('size-a4', 'size-letter');
      wordPage.classList.add(`size-${size}`);
      showToast(`Đã đổi cỡ giấy: ${btn.textContent.trim()}`);
    });
  });

  // View Tab Actions
  document.getElementById('toggleRuler').addEventListener('change', (e) => {
    rulerContainer.style.display = e.target.checked ? 'flex' : 'none';
  });

  document.getElementById('btnZoom100').addEventListener('click', () => setZoom(100));
  document.getElementById('btnZoomFit').addEventListener('click', () => {
    const workspaceWidth = document.getElementById('wordWorkspace').clientWidth - 60;
    const pageWidth = wordPage.offsetWidth;
    const fitPercent = Math.round((workspaceWidth / pageWidth) * 100);
    setZoom(fitPercent);
  });

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen().catch(() => {});
    }
  };
  document.getElementById('btnToggleFullscreen').addEventListener('click', toggleFullscreen);
  document.getElementById('sbToggleFullscreen').addEventListener('click', toggleFullscreen);

  // Zoom Controls in Statusbar
  zoomSlider.addEventListener('input', (e) => setZoom(parseInt(e.target.value, 10)));
  document.getElementById('btnZoomIn').addEventListener('click', () => setZoom(state.zoomLevel + 10));
  document.getElementById('btnZoomOut').addEventListener('click', () => setZoom(state.zoomLevel - 10));

  // Document Title Change
  docTitleInput.addEventListener('input', (e) => {
    state.currentFileName = e.target.value.trim() || 'TaiLieu';
    document.getElementById('bsDocName').textContent = `${state.currentFileName}.docx`;
    setSaveStatus('unsaved');
  });

  // Drag and drop file onto workspace
  const workspace = document.getElementById('wordWorkspace');
  ['dragenter', 'dragover'].forEach(eventName => {
    workspace.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      workspace.style.backgroundColor = '#deecf9';
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    workspace.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      workspace.style.backgroundColor = '';
    }, false);
  });

  workspace.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const file = dt.files[0];
    if (file) {
      handleFileImport(file, 'auto');
    }
  });

  // Dropzone in Backstage
  const dropzone = document.getElementById('backstageDropzone');
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) {
      handleFileImport(file, 'auto');
      backstageModal.style.display = 'none';
    }
  });

  // Keyboard Shortcuts
  window.addEventListener('keydown', (e) => {
    // Ctrl + S or Cmd + S -> Save DOCX
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      saveDocument();
    }
    // Ctrl + P or Cmd + P -> Print
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
      e.preventDefault();
      window.print();
    }
    // Ctrl + O or Cmd + O -> Open file
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
      e.preventDefault();
      docxFileInput.click();
    }
    // Escape -> Close Backstage
    if (e.key === 'Escape' && backstageModal.style.display !== 'none') {
      backstageModal.style.display = 'none';
    }
  });
}

/**
 * Main Initialization
 */
async function main() {
  initIcons();
  renderRulerMarks();
  setupEventListeners();

  // Initialize CKEditor 5 Decoupled Editor
  const editorElement = document.getElementById('editor');
  const toolbarContainer = document.getElementById('ckeditorToolbarContainer');

  try {
    state.editor = await initWordEditor(
      editorElement,
      toolbarContainer,
      ({ words, characters }) => {
        sbWordCount.textContent = `${words} từ`;
        sbCharCount.textContent = `${characters} ký tự`;
        document.getElementById('bsWordCount').textContent = `${words} từ (${characters} ký tự)`;
      }
    );
    window.editor = state.editor;
    window.DocxService = DocxService;
    window.appState = state;

    // Sync line spacing ribbon button with cursor position in document
    state.editor.model.document.selection.on('change:range', () => {
      const firstBlock = Array.from(state.editor.model.document.selection.getSelectedBlocks())[0];
      if (firstBlock) {
        const activeLh = firstBlock.getAttribute('lineHeight') || '1.15';
        const currentLineSpacingText = document.getElementById('currentLineSpacingText');
        if (currentLineSpacingText && currentLineSpacingText.textContent !== activeLh) {
          currentLineSpacingText.textContent = activeLh;
          document.querySelectorAll('[data-lineheight]').forEach(i => {
            i.classList.toggle('active', i.dataset.lineheight === activeLh);
          });
        }
      }
    });

    // Track unsaved changes & debounce auto-save
    state.editor.model.document.on('change:data', () => {
      setSaveStatus('unsaved');
      clearTimeout(state.autoSaveTimer);
      state.autoSaveTimer = setTimeout(() => {
        if (state.editor) {
          DocxService.saveToLocalStorage(state.editor.getData(), docTitleInput.value.trim());
          setSaveStatus('saved');
        }
      }, 2500);
    });

    // Check for previous draft in LocalStorage
    const draft = DocxService.loadFromLocalStorage();
    if (draft && draft.content) {
      console.log('Phát hiện bản thảo đã lưu từ phiên trước.');
    }

    sbMessage.textContent = 'Sẵn sàng soạn thảo';
  } catch (error) {
    console.error('Không thể khởi tạo CKEditor 5:', error);
    showToast(`Lỗi khởi tạo trình soạn thảo: ${error.message}`, true);
  }
}

// Bootstrap
window.addEventListener('DOMContentLoaded', main);
