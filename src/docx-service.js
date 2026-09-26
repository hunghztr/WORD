import { renderAsync } from 'docx-preview';
import mammoth from 'mammoth';
import { asBlob } from 'html-docx-js-typescript';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

/**
 * Service to handle DOCX preview, import, export, and auto-save
 */
export class DocxService {
  /**
   * Detect dominant font family and font size from DOCX OOXML structures
   * @param {ArrayBuffer} arrayBuffer 
   * @returns {Promise<{ font: string, fontSize: number }>}
   */
  static async detectDocxFontProperties(arrayBuffer) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const fontCounts = {};
      const sizeCounts = {};
      const lineSpacingCounts = {};
      const alignCounts = {};

      for (const [filePath, file] of Object.entries(zip.files)) {
        if (filePath.startsWith('word/') && !file.dir) {
          const content = await file.async('text');

          // 1. Extract font names from w:rFonts (Word XML)
          const fontRegex = /w:rFonts[^>]+w:(?:ascii|hAnsi|cs)="([^"]+)"/g;
          let match;
          while ((match = fontRegex.exec(content)) !== null) {
            const fontName = match[1].trim();
            if (fontName && !['Symbol', 'Wingdings', 'Webdings'].includes(fontName)) {
              fontCounts[fontName] = (fontCounts[fontName] || 0) + 1;
            }
          }

          // 2. Check w:font w:name in fontTable.xml
          const fontTableRegex = /w:font[^>]+w:name="([^"]+)"/g;
          while ((match = fontTableRegex.exec(content)) !== null) {
            const fontName = match[1].trim();
            if (fontName && !['Symbol', 'Wingdings', 'Webdings'].includes(fontName)) {
              fontCounts[fontName] = (fontCounts[fontName] || 0) + 1;
            }
          }

          // 3. Extract CSS font-family (altChunk / MHT / HTML)
          const cssFontRegex = /font-family:\s*['"]?([^;'"\r\n]+)['"]?/gi;
          while ((match = cssFontRegex.exec(content)) !== null) {
            const fontName = match[1].split(',')[0].replace(/['"]/g, '').trim();
            if (fontName && !['serif', 'sans-serif', 'monospace'].includes(fontName.toLowerCase())) {
              fontCounts[fontName] = (fontCounts[fontName] || 0) + 1;
            }
          }

          // 4. Extract font sizes from w:sz w:val (Word XML half-points)
          const sizeRegex = /w:sz[^>]+w:val="(\d+)"/g;
          while ((match = sizeRegex.exec(content)) !== null) {
            const val = parseInt(match[1], 10);
            const pt = Math.round(val / 2);
            if (pt >= 8 && pt <= 72) {
              sizeCounts[pt] = (sizeCounts[pt] || 0) + 1;
            }
          }

          // 5. Extract CSS font-size (pt / px)
          const cssSizeRegex = /font-size:\s*(\d+(?:\.\d+)?)\s*(?:pt|px)?/gi;
          while ((match = cssSizeRegex.exec(content)) !== null) {
            const pt = Math.round(parseFloat(match[1]));
            if (pt >= 8 && pt <= 72) {
              sizeCounts[pt] = (sizeCounts[pt] || 0) + 1;
            }
          }

          // 6. Extract line spacing from w:spacing w:line
          const lineRegex = /w:spacing[^>]+w:line="(\d+)"/g;
          while ((match = lineRegex.exec(content)) !== null) {
            const val = parseInt(match[1], 10);
            if (val > 0) {
              let lh = (val / 240).toFixed(2);
              if (lh.endsWith('.00')) lh = lh.slice(0, -3);
              else if (lh.endsWith('0')) lh = lh.slice(0, -1);
              lineSpacingCounts[lh] = (lineSpacingCounts[lh] || 0) + 1;
            }
          }

          // 7. Extract paragraph alignment from w:jc
          const jcRegex = /w:jc[^>]+w:val="([^"]+)"/g;
          while ((match = jcRegex.exec(content)) !== null) {
            let val = match[1].toLowerCase().trim();
            if (val === 'both' || val === 'distribute') val = 'justify';
            if (['left', 'center', 'right', 'justify'].includes(val)) {
              alignCounts[val] = (alignCounts[val] || 0) + 1;
            }
          }
        }
      }

      // Find dominant font
      let primaryFont = 'Calibri';
      let maxFontCount = 0;
      for (const [font, count] of Object.entries(fontCounts)) {
        if (count > maxFontCount) {
          maxFontCount = count;
          primaryFont = font;
        }
      }

      // If document is standard Vietnamese doc (often Times New Roman)
      if (fontCounts['Times New Roman']) {
        primaryFont = 'Times New Roman';
      }

      // Find dominant size
      let primarySize = 14;
      let maxSizeCount = 0;
      for (const [sz, count] of Object.entries(sizeCounts)) {
        if (count > maxSizeCount) {
          maxSizeCount = count;
          primarySize = parseInt(sz, 10);
        }
      }

      // Find dominant line spacing
      let primaryLineSpacing = '1.15';
      let maxLhCount = 0;
      for (const [lh, count] of Object.entries(lineSpacingCounts)) {
        if (count > maxLhCount) {
          maxLhCount = count;
          primaryLineSpacing = lh;
        }
      }

      // Find dominant alignment
      let primaryAlign = 'left';
      let maxAlignCount = 0;
      for (const [al, count] of Object.entries(alignCounts)) {
        if (count > maxAlignCount) {
          maxAlignCount = count;
          primaryAlign = al;
        }
      }

      return {
        font: primaryFont,
        fontSize: primarySize,
        lineSpacing: primaryLineSpacing,
        alignment: primaryAlign
      };
    } catch (e) {
      console.warn('Không thể phân tích font từ DOCX:', e);
      return { font: 'Times New Roman', fontSize: 14, lineSpacing: '1.15', alignment: 'left' };
    }
  }

  /**
   * Detect page size and margins from DOCX document.xml
   * @param {ArrayBuffer} arrayBuffer
   * @returns {Promise<{ pageWidthPx: number, pageHeightPx: number, marginTopPx: number, marginRightPx: number, marginBottomPx: number, marginLeftPx: number, orientation: string }>}
   */
  static async detectDocxPageLayout(arrayBuffer) {
    // TWIPs -> px at 96 DPI: 1 inch = 1440 TWIPs, 1 inch = 96 px => 1 TWIPs = 96/1440 px
    const TWIPS_TO_PX = 96 / 1440;
    const defaults = {
      pageWidthPx: 794,   // A4 210mm at 96 DPI
      pageHeightPx: 1123, // A4 297mm at 96 DPI
      marginTopPx: 96,
      marginRightPx: 96,
      marginBottomPx: 96,
      marginLeftPx: 96,
      orientation: 'portrait'
    };

    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const docFile = zip.file('word/document.xml');
      if (!docFile) return defaults;

      const xml = await docFile.async('text');

      // Read w:pgSz (use dotAll flag for robustness)
      const pgSzTagMatch = xml.match(/<w:pgSz\s[^>]*?\/?>/s);
      let wW = 12240, wH = 15840; // default A4 in TWIPs
      if (pgSzTagMatch) {
        const tag = pgSzTagMatch[0];
        const wAttr = tag.match(/w:w="(\d+)"/);
        const hAttr = tag.match(/w:h="(\d+)"/);
        if (wAttr) wW = parseInt(wAttr[1], 10);
        if (hAttr) wH = parseInt(hAttr[1], 10);
      }

      // Read w:pgMar (use dotAll flag for robustness)
      const pgMarTagMatch = xml.match(/<w:pgMar\s[^>]*?\/?>/s);
      let mTop = 1440, mRight = 1440, mBottom = 1440, mLeft = 1440;
      if (pgMarTagMatch) {
        const tag = pgMarTagMatch[0];
        const topA = tag.match(/w:top="(\d+)"/);
        const rightA = tag.match(/w:right="(\d+)"/);
        const bottomA = tag.match(/w:bottom="(\d+)"/);
        const leftA = tag.match(/w:left="(\d+)"/);
        if (topA) mTop = parseInt(topA[1], 10);
        if (rightA) mRight = parseInt(rightA[1], 10);
        if (bottomA) mBottom = parseInt(bottomA[1], 10);
        if (leftA) mLeft = parseInt(leftA[1], 10);
      }

      const pageWidthPx = Math.round(wW * TWIPS_TO_PX);
      const pageHeightPx = Math.round(wH * TWIPS_TO_PX);
      const orientation = pageWidthPx > pageHeightPx ? 'landscape' : 'portrait';

      return {
        pageWidthPx,
        pageHeightPx,
        marginTopPx: Math.round(mTop * TWIPS_TO_PX),
        marginRightPx: Math.round(mRight * TWIPS_TO_PX),
        marginBottomPx: Math.round(mBottom * TWIPS_TO_PX),
        marginLeftPx: Math.round(mLeft * TWIPS_TO_PX),
        orientation
      };
    } catch (e) {
      console.warn('Không thể đọc page layout từ DOCX:', e);
      return defaults;
    }
  }

  /**
   * Build a per-paragraph style map from document.xml
   * Returns an array of { lineHeight, alignment } for each <w:p> in document order.
   * @param {ArrayBuffer} arrayBuffer
   * @returns {Promise<Array<{lineHeight: string|null, alignment: string|null}>>}
   */
  static async buildParagraphStyleMap(arrayBuffer) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);

      // 1. Read styles.xml to build named style -> {lineHeight, alignment} map
      const styleDefaults = {};
      const stylesFile = zip.file('word/styles.xml');
      if (stylesFile) {
        const stylesXml = await stylesFile.async('text');
        // Match each <w:style> block
        const styleBlockReg = /<w:style\s[^>]*?>([\/\s\S]*?)<\/w:style>/g;
        let sm;
        while ((sm = styleBlockReg.exec(stylesXml)) !== null) {
          const block = sm[0];
          const idMatch = block.match(/w:styleId="([^"]+)"/);
          if (!idMatch) continue;
          const styleId = idMatch[1];
          const lhRaw = block.match(/w:spacing[^>]*?w:line="(\d+)"/);
          const jcRaw = block.match(/w:jc[^>]*?w:val="([^"]+)"/);
          styleDefaults[styleId] = {
            lineHeight: lhRaw ? String(Math.round((parseInt(lhRaw[1], 10) / 240) * 100) / 100) : null,
            alignment: jcRaw ? ((['both','distribute'].includes(jcRaw[1]) ? 'justify' : jcRaw[1])) : null
          };
        }
      }

      // 2. Read document.xml paragraph by paragraph
      const docFile = zip.file('word/document.xml');
      if (!docFile) return [];
      const xml = await docFile.async('text');

      // Split into w:p blocks
      const paraBlocks = [];
      const paraReg = /<w:p[\s>][\/\s\S]*?<\/w:p>/g;
      let pm;
      while ((pm = paraReg.exec(xml)) !== null) {
        paraBlocks.push(pm[0]);
      }

      return paraBlocks.map(para => {
        // Find referenced style
        const styleMatch = para.match(/<w:pStyle[^>]+w:val="([^"]+)"/);
        const styleId = styleMatch ? styleMatch[1] : null;
        const styleBase = (styleId && styleDefaults[styleId]) ? styleDefaults[styleId] : {};

        // Direct w:spacing on this paragraph overrides style
        const spacingMatch = para.match(/<w:spacing[^>]*?w:line="(\d+)"[^>]*?(?:\/>|>)/);
        let lineHeight = styleBase.lineHeight || null;
        if (spacingMatch) {
          const rawVal = parseInt(spacingMatch[1], 10);
          lineHeight = String(Math.round((rawVal / 240) * 100) / 100);
        }

        // Direct w:jc on this paragraph overrides style
        const jcMatch = para.match(/<w:jc[^>]*?w:val="([^"]+)"/);
        let alignment = styleBase.alignment || null;
        if (jcMatch) {
          let jcVal = jcMatch[1].toLowerCase();
          if (jcVal === 'both' || jcVal === 'distribute') jcVal = 'justify';
          if (['left', 'center', 'right', 'justify'].includes(jcVal)) {
            alignment = jcVal;
          }
        }

        return { lineHeight, alignment };
      });
    } catch (e) {
      console.warn('buildParagraphStyleMap error:', e);
      return [];
    }
  }

  /**
   * Decode quoted-printable string
   * @param {string} str 
   * @returns {string}
   */
  static decodeQuotedPrintable(str) {
    return str
      .replace(/=\r?\n/g, '')
      .replace(/=([0-9A-Fa-f]{2})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
  }

  /**
   * Extract HTML from MHTML format (word/afchunk.mht)
   * @param {string} mhtText 
   * @returns {string}
   */
  static extractHtmlFromMht(mhtText) {
    const decoded = DocxService.decodeQuotedPrintable(mhtText);
    const bodyMatch = decoded.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    if (bodyMatch) return bodyMatch[1].trim();
    const htmlMatch = decoded.match(/<html[^>]*>([\s\S]*?)<\/html>/i);
    if (htmlMatch) return htmlMatch[1].trim();

    const parts = decoded.split(/\r?\n\r?\n/);
    for (let i = 1; i < parts.length; i++) {
      const chunk = parts[i].trim();
      if (chunk.startsWith('<') && !chunk.startsWith('------')) {
        return chunk.split('------=')[0].trim();
      }
    }
    return decoded;
  }

  /**
   * Copy the computed (class/stylesheet-derived) formatting of a docx-preview render
   * into inline styles, so it survives once classes and <style> tags are dropped.
   * Must be called while the render is attached to the document.
   * @param {HTMLElement} root
   */
  static inlineComputedStyles(root) {
    const toPt = px => `${Math.round(parseFloat(px) * 0.75 * 100) / 100}pt`;
    const isZero = v => !v || parseFloat(v) === 0;
    const isTransparent = c => !c || c === 'transparent' || c === 'rgba(0, 0, 0, 0)';

    // Paragraph-level formatting
    root.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li').forEach(el => {
      const cs = getComputedStyle(el);
      el.style.marginTop = toPt(cs.marginTop);
      el.style.marginBottom = toPt(cs.marginBottom);
      for (const prop of ['marginLeft', 'marginRight', 'paddingLeft', 'paddingRight', 'textIndent']) {
        if (!isZero(cs[prop])) el.style[prop] = toPt(cs[prop]);
      }

      let align = cs.textAlign;
      if (align === 'start') align = 'left';
      if (align === 'end') align = 'right';
      if (['left', 'center', 'right', 'justify'].includes(align)) el.style.textAlign = align;

      // Line height: "auto" rule is a multiple (unitless, scales with font size),
      // "exact"/"atLeast" are absolute. Probe by changing font-size to tell them apart.
      if (cs.lineHeight !== 'normal') {
        const lhPx = parseFloat(cs.lineHeight);
        const prevFontSize = el.style.fontSize;
        el.style.fontSize = '100px';
        const probePx = parseFloat(getComputedStyle(el).lineHeight);
        el.style.fontSize = prevFontSize;
        if (Math.abs(probePx - lhPx) > 0.5) {
          const ratio = Math.round((probePx / 100) * 100) / 100;
          el.style.lineHeight = String(ratio);
          el.setAttribute('data-lineheight', String(ratio));
        } else {
          el.style.lineHeight = toPt(lhPx);
        }
      }
    });

    // Run-level formatting
    root.querySelectorAll('span').forEach(el => {
      if (!el.textContent) return;
      const cs = getComputedStyle(el);
      el.style.fontFamily = cs.fontFamily;
      el.style.fontSize = toPt(cs.fontSize);
      if (parseInt(cs.fontWeight, 10) >= 600) el.style.fontWeight = 'bold';
      if (cs.fontStyle === 'italic') el.style.fontStyle = 'italic';
      const deco = cs.textDecorationLine;
      if (deco && deco !== 'none') el.style.textDecoration = deco;
      if (cs.color !== 'rgb(0, 0, 0)') el.style.color = cs.color;
      if (!isTransparent(cs.backgroundColor)) el.style.backgroundColor = cs.backgroundColor;
      if (cs.textTransform !== 'none') el.style.textTransform = cs.textTransform;
      if (cs.verticalAlign === 'super' || cs.verticalAlign === 'sub') el.style.verticalAlign = cs.verticalAlign;
      if (cs.letterSpacing !== 'normal' && !isZero(cs.letterSpacing)) el.style.letterSpacing = toPt(cs.letterSpacing);
    });

    // Table / cell formatting (borders, padding, shading) from Word table styles
    root.querySelectorAll('table').forEach(tbl => {
      const cs = getComputedStyle(tbl);
      if (!isZero(cs.marginLeft)) tbl.style.marginLeft = toPt(cs.marginLeft);

      // Word grid column widths come as <colgroup>, which CKEditor drops:
      // move them onto the cells (cell width is kept by TableCellProperties)
      const colWidths = [...tbl.querySelectorAll(':scope > colgroup > col')].map(col => parseFloat(col.style.width) || 0);
      if (!colWidths.length) return;
      for (const row of tbl.rows) {
        let colIndex = 0;
        for (const cell of row.cells) {
          const span = cell.colSpan || 1;
          const gridPt = colWidths.slice(colIndex, colIndex + span).reduce((a, b) => a + b, 0);
          const ccs = getComputedStyle(cell);
          const paddingPt = (parseFloat(ccs.paddingLeft) + parseFloat(ccs.paddingRight)) * 0.75;
          const widthPt = Math.round((gridPt - paddingPt) * 100) / 100;
          if (widthPt > 0 && !cell.style.width) cell.style.width = `${widthPt}pt`;
          colIndex += span;
        }
      }
    });
    root.querySelectorAll('td, th').forEach(cell => {
      const cs = getComputedStyle(cell);
      for (const side of ['Top', 'Right', 'Bottom', 'Left']) {
        cell.style[`padding${side}`] = toPt(cs[`padding${side}`]);
        const style = cs[`border${side}Style`];
        if (style === 'none' || style === 'hidden' || isZero(cs[`border${side}Width`])) {
          cell.style[`border${side}`] = 'none';
        } else {
          cell.style[`border${side}`] = `${toPt(cs[`border${side}Width`])} ${style} ${cs[`border${side}Color`]}`;
        }
      }
      // Row banding shading is applied on <tr>; move it onto the cell for CKEditor
      const bg = isTransparent(cs.backgroundColor) ? getComputedStyle(cell.parentElement).backgroundColor : cs.backgroundColor;
      if (!isTransparent(bg)) cell.style.backgroundColor = bg;
      cell.style.verticalAlign = cs.verticalAlign === 'baseline' ? 'top' : cs.verticalAlign;
    });
  }

  /**
   * Render DOCX with docx-preview (offscreen) and extract clean HTML for CKEditor.
   * This gives maximum fidelity: all Word formatting (font, size, bold, alignment,
   * line-height, tables, borders) is correctly read by docx-preview.
   * @param {ArrayBuffer} arrayBuffer
   * @returns {Promise<string>} Clean HTML string ready for CKEditor
   */
  static async extractEditorHtmlFromDocxPreview(arrayBuffer) {
    // Create a hidden offscreen container
    const container = document.createElement('div');
    container.style.cssText = 'position:absolute;left:-99999px;top:-99999px;width:794px;visibility:hidden;pointer-events:none;';
    document.body.appendChild(container);

    try {
      await renderAsync(arrayBuffer, container, null, {
        className: 'docx-preview-rendered',
        inWrapper: false,       // No wrapper — gives us raw section content
        ignoreWidth: true,
        ignoreHeight: true,
        ignoreFonts: false,
        breakPages: false,      // Single continuous flow for editor
        renderChanges: false,
        renderHeaders: false,   // Skip header/footer (editor handles separately)
        renderFooters: false,
        renderAltChunks: true
      });

      // docx-preview puts most Word formatting (docDefaults, paragraph/table styles)
      // in <style> class rules. Bake the computed values into inline styles while the
      // offscreen render is still attached, before classes and <style> are stripped.
      DocxService.inlineComputedStyles(container);

      // Floating shapes (drawn lines etc.) are absolutely positioned SVG inside <div>s
      // nested in a <p>. That is invalid HTML: re-parsing would split the paragraph into
      // several empty blocks, and CKEditor cannot keep the shapes anyway. Drop them here.
      container.querySelectorAll('svg, canvas').forEach(el => el.remove());
      container.querySelectorAll('p div').forEach(el => {
        if (!el.textContent.trim() && !el.querySelector('img')) el.remove();
      });

      // Word keeps consecutive spaces (used to position text); HTML/CKEditor collapses
      // them. Turn each run into non-breaking spaces, keeping one normal space to wrap on.
      const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        if (node.parentElement?.tagName !== 'STYLE' && / {2,}/.test(node.data)) {
          node.data = node.data.replace(/ {2,}/g, run => ' '.repeat(run.length - 1) + ' ');
        }
      }

      // docx-preview renders sections: .docx-preview-section elements
      // Extract all content and inline the critical styles
      const sections = container.querySelectorAll('section.docx-wrapper, .docx-preview-section, section, article');
      let htmlParts = [];

      if (sections.length > 0) {
        sections.forEach(section => {
          htmlParts.push(section.innerHTML);
        });
      } else {
        // Fallback: take everything in the container
        htmlParts.push(container.innerHTML);
      }

      const rawHtml = htmlParts.join('\n');

      // Parse and normalize for CKEditor
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = rawHtml;

      // 1. docx-preview uses <span class="docx-..."> wrappers — preserve inline styles
      // Remove docx-specific class prefixes but keep inline styles
      tempDiv.querySelectorAll('[class]').forEach(el => {
        const cls = el.getAttribute('class') || '';
        if (cls.startsWith('docx-')) {
          el.removeAttribute('class');
        }
      });

      // 2. Convert <span style="font-weight:bold"> → preserve (CKEditor handles inline styles via GHS)
      // 3. Tables: borders/padding were already inlined from the real Word table style
      tempDiv.querySelectorAll('table').forEach(tbl => {
        tbl.style.borderCollapse = 'collapse';
      });

      // 4. Remove empty paragraphs that are just spacers (keep at most 1 consecutive)

      return tempDiv.innerHTML;
    } finally {
      // Always clean up the offscreen container
      if (container.parentNode) container.parentNode.removeChild(container);
    }
  }

  /**
   * Preview a DOCX file buffer into a DOM container using docx-preview

   * @param {ArrayBuffer|Blob|File} fileData 
   * @param {HTMLElement} container 
   */
  static async previewDocx(fileData, container) {
    container.innerHTML = '<div class="preview-loading"><div class="spinner"></div><p>Đang kết xuất tài liệu DOCX chuẩn Word...</p></div>';
    try {
      let arrayBuffer;
      if (fileData instanceof ArrayBuffer) {
        arrayBuffer = fileData;
      } else if (fileData instanceof Blob || fileData instanceof File) {
        arrayBuffer = await fileData.arrayBuffer();
      } else {
        throw new Error('Dữ liệu file không hợp lệ.');
      }

      container.innerHTML = '';
      await renderAsync(arrayBuffer, container, null, {
        className: 'docx-preview-rendered',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        renderChanges: false,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        renderAltChunks: true
      });
      return { success: true };
    } catch (error) {
      console.error('Lỗi khi render docx-preview:', error);
      container.innerHTML = `
        <div class="preview-error">
          <span class="error-icon">⚠️</span>
          <h3>Không thể xem trước tệp DOCX này</h3>
          <p>${error.message || 'Định dạng tệp không được hỗ trợ hoặc tệp bị hỏng.'}</p>
        </div>
      `;
      return { success: false, error };
    }
  }

  /**
   * Import DOCX or text/html file and convert to clean HTML for CKEditor
   * @param {File} file 
   * @returns {Promise<{ html: string, title: string, rawBuffer: ArrayBuffer, font: string, fontSize: number }>}
   */
  static async importDocFile(file) {
    const fileName = file.name;
    const isDocx = fileName.toLowerCase().endsWith('.docx');
    const isDoc = fileName.toLowerCase().endsWith('.doc');

    const arrayBuffer = await file.arrayBuffer();

    if (isDocx) {
      let rawHtml = '';
      let warnings = [];
      let fromDocxPreview = false;

      // 2. Detect page layout (size + margins) from document XML
      const fontProps = await DocxService.detectDocxFontProperties(arrayBuffer);
      const pageLayout = await DocxService.detectDocxPageLayout(arrayBuffer);

      // STRATEGY: Use docx-preview (renders 100% of Word formatting) as PRIMARY source.
      // Fall back to MHT extraction, then mammoth if docx-preview fails.

      // Primary: docx-preview offscreen render → extract HTML
      try {
        rawHtml = await DocxService.extractEditorHtmlFromDocxPreview(arrayBuffer);
        if (rawHtml && rawHtml.trim().length > 20) {
          fromDocxPreview = true;
          console.info('[Import] Using docx-preview extraction (high fidelity)');
        } else {
          rawHtml = '';
        }
      } catch (previewErr) {
        console.warn('[Import] docx-preview extraction failed, falling back:', previewErr);
        rawHtml = '';
      }

      // Fallback 1: MHT chunk (for html-docx-js exported files)
      if (!rawHtml) {
        try {
          const zip = await JSZip.loadAsync(arrayBuffer);
          const mhtFile = zip.file('word/afchunk.mht');
          if (mhtFile) {
            const mhtText = await mhtFile.async('text');
            rawHtml = DocxService.extractHtmlFromMht(mhtText);
            console.info('[Import] Using MHT extraction');
          }
        } catch (zipErr) {
          console.warn('Cannot check docx zip for afchunk:', zipErr);
        }
      }

      // Fallback 2: mammoth (basic OOXML conversion)
      if (!rawHtml) {
        try {
          const result = await mammoth.convertToHtml(
            { arrayBuffer: arrayBuffer },
            {
              styleMap: [
                "p[style-name='Heading 1'] => h1:fresh",
                "p[style-name='Heading 2'] => h2:fresh",
                "p[style-name='Heading 3'] => h3:fresh",
                "p[style-name='Title'] => h1.title:fresh",
                "p[style-name='Subtitle'] => p.subtitle:fresh"
              ]
            }
          );
          rawHtml = result.value || '<p>Tài liệu trống.</p>';
          warnings = result.messages;
          console.info('[Import] Using mammoth fallback');
        } catch (mamErr) {
          rawHtml = '<p>Không thể đọc nội dung tài liệu.</p>';
        }
      }

      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = rawHtml;

      // Standardize table formatting & cell attributes for CKEditor 5
      // (docx-preview output already carries the real Word table borders/padding)
      const tables = fromDocxPreview ? [] : tempDiv.querySelectorAll('table');
      tables.forEach(table => {
        table.style.borderCollapse = 'collapse';
        if (!table.style.width) table.style.width = '100%';
        if (!table.style.margin) table.style.margin = '14px 0';
        if (!table.style.border) table.style.border = '1px solid #bfbfbf';

        table.querySelectorAll('th').forEach(th => {
          if (!th.style.backgroundColor) th.style.backgroundColor = '#f2f2f2';
          if (!th.style.border) th.style.border = '1px solid #bfbfbf';
          if (!th.style.padding) th.style.padding = '6px 10px';
          if (!th.style.fontWeight) th.style.fontWeight = 'bold';
        });

        table.querySelectorAll('td').forEach(td => {
          if (!td.style.border) td.style.border = '1px solid #bfbfbf';
          if (!td.style.padding) td.style.padding = '6px 10px';
          if (!td.style.verticalAlign) td.style.verticalAlign = 'top';
        });
      });

      // Ensure cells without block tags wrap their text in a paragraph
      tempDiv.querySelectorAll('td, th').forEach(cell => {
        if (!cell.querySelector('p, h1, h2, h3, ul, ol') && cell.textContent.trim()) {
          const directAlign = cell.getAttribute('align') || cell.style.textAlign;
          const alignAttr = directAlign ? ` align="${directAlign}" class="text-align-${directAlign}" style="text-align: ${directAlign};"` : '';
          cell.innerHTML = `<p${alignAttr}>${cell.innerHTML}</p>`;
        }
      });

      // Determine if source is high-fidelity (docx-preview) or low-fidelity (mammoth)
      // docx-preview output has many inline style spans with font-family already set
      const isHighFidelity = fromDocxPreview || rawHtml.includes('font-family') || rawHtml.includes('font-size');

      // Build per-paragraph style map for mammoth fallback alignment/lineheight
      const paraStyleMap = isHighFidelity ? [] : await DocxService.buildParagraphStyleMap(arrayBuffer);

      const blocks = tempDiv.querySelectorAll('p, h1, h2, h3, li');
      blocks.forEach((block, idx) => {
        const paraStyle = paraStyleMap[idx] || {};

        // --- Alignment ---
        let align = block.style.textAlign && block.style.textAlign !== 'initial'
          ? block.style.textAlign
          : block.getAttribute('align') || paraStyle.alignment || null;
        if (align) {
          block.style.textAlign = align;
          block.setAttribute('align', align);
          ['left','center','right','justify'].forEach(a => block.classList.remove(`text-align-${a}`));
          block.classList.add(`text-align-${align}`);
        }

        // --- Line Height ---
        // Always normalize to data-lineheight for CKEditor plugin to pick up
        let lh = block.getAttribute('data-lineheight')
          || (block.style.lineHeight && block.style.lineHeight !== 'normal' ? block.style.lineHeight : null)
          || paraStyle.lineHeight
          || (isHighFidelity ? null : fontProps.lineSpacing)
          || null;
        if (lh) {
          block.style.lineHeight = lh;
          block.setAttribute('data-lineheight', lh);
        }

        // --- Font wrap: ONLY for mammoth fallback (docx-preview already has inline styles) ---
        if (!isHighFidelity) {
          const spansWithFont = block.querySelectorAll('span[style*="font-family"]');
          if (spansWithFont.length === 0 && block.textContent.trim()) {
            const inner = block.innerHTML;
            block.innerHTML = `<span style="font-family: '${fontProps.font}', serif; font-size: ${fontProps.fontSize}pt;">${inner}</span>`;
          }
        }
      });

      const cleanName = fileName.replace(/\.[^/.]+$/, '');

      return {

        html: tempDiv.innerHTML,
        title: cleanName,
        rawBuffer: arrayBuffer,
        font: fontProps.font,
        fontSize: fontProps.fontSize,
        lineSpacing: fontProps.lineSpacing,
        alignment: fontProps.alignment,
        pageLayout: pageLayout,
        warnings: warnings
      };
    } else if (isDoc) {
      // Legacy .doc (Word 97-2003 binary format)
      // Attempt to extract text or notify user
      try {
        const textDecoder = new TextDecoder('utf-8');
        const textContent = textDecoder.decode(arrayBuffer);
        
        // If it happens to be RTF or HTML saved as .doc
        if (textContent.includes('<html') || textContent.includes('{\\rtf')) {
          let html = textContent;
          if (textContent.includes('{\\rtf')) {
            // Basic RTF fallback: strip common rtf controls
            const stripped = textContent.replace(/\\par[d]?/g, '<br>').replace(/\\[a-z0-9\-]+/gi, '').replace(/[{}]/g, '');
            html = `<p>${stripped}</p>`;
          }
          return {
            html,
            title: fileName.replace(/\.[^/.]+$/, ''),
            rawBuffer: arrayBuffer,
            isLegacyDoc: true
          };
        }

        // Try mammoth in case it is actually docx renamed to .doc
        try {
          const result = await mammoth.convertToHtml({ arrayBuffer });
          if (result.value) {
            return {
              html: result.value,
              title: fileName.replace(/\.[^/.]+$/, ''),
              rawBuffer: arrayBuffer,
              isLegacyDoc: false
            };
          }
        } catch (_) {}

        // Fallback for legacy binary DOC
        throw new Error('Tệp .doc (Word 97-2003 nhị phân) là định dạng cũ. Vui lòng mở bằng Microsoft Word và "Save As" sang định dạng .docx (Word 2007+) để giữ đầy đủ định dạng và hình ảnh!');
      } catch (err) {
        throw err;
      }
    } else {
      throw new Error('Chỉ hỗ trợ tệp định dạng .docx hoặc .doc');
    }
  }

  /**
   * Export HTML string to a standard Microsoft Word .docx file
   * @param {string} htmlContent 
   * @param {string} fileName 
   * @param {object} options 
   */
  static async exportToDocx(htmlContent, fileName = 'TaiLieu.docx', options = {}) {
    const documentTitle = fileName.replace(/\.[^/.]+$/, '');
    
    // Pre-process HTML: unwrap figure.table, enforce borders and inline alignment
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlContent;

    // 1. Unwrap CKEditor <figure class="table"> so Word renders plain <table>
    tempDiv.querySelectorAll('figure.table').forEach(fig => {
      const table = fig.querySelector('table');
      if (table) {
        fig.parentNode.replaceChild(table, fig);
      }
    });

    // 2. Ensure all tables have border and collapse styling
    tempDiv.querySelectorAll('table').forEach(tbl => {
      tbl.setAttribute('border', '1');
      tbl.style.borderCollapse = 'collapse';
      if (!tbl.style.width) tbl.style.width = '100%';
      tbl.style.border = '1px solid #bfbfbf';
    });

    // 3. Ensure cells have border, padding, and vertical alignment
    tempDiv.querySelectorAll('th, td').forEach(cell => {
      cell.style.border = '1px solid #bfbfbf';
      if (!cell.style.padding) cell.style.padding = '6pt 10pt';
      if (!cell.style.verticalAlign) cell.style.verticalAlign = 'top';
      if (cell.tagName === 'TH' && !cell.style.backgroundColor) {
        cell.style.backgroundColor = '#f2f2f2';
      }
    });

    // 4. Ensure paragraph alignment and line-height are explicitly in style & attributes
    tempDiv.querySelectorAll('p, h1, h2, h3, li, td, th').forEach(el => {
      let align = el.style.textAlign || el.getAttribute('align');
      if (!align) {
        for (const al of ['center', 'right', 'justify', 'left']) {
          if (el.classList.contains(`text-align-${al}`)) {
            align = al;
            break;
          }
        }
      }
      if (align) {
        el.style.textAlign = align;
        el.setAttribute('align', align);
      }

      const lh = el.getAttribute('data-lineheight') || el.style.lineHeight;
      if (lh) {
        el.style.lineHeight = lh;
      }
    });

    const cleanHtml = tempDiv.innerHTML;

    // Wrap with full HTML structure and standard Word CSS styling
    const fullHtml = `
      <!DOCTYPE html>
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset="utf-8">
        <title>${documentTitle}</title>
        <style>
          @page {
            size: A4 portrait;
            margin: 25.4mm 25.4mm 25.4mm 25.4mm;
            mso-header-margin: 12.7mm;
            mso-footer-margin: 12.7mm;
          }
          body {
            font-family: 'Calibri', 'Times New Roman', 'Segoe UI', Arial, sans-serif;
            font-size: 11pt;
            line-height: 1.25;
            color: #000000;
          }
          h1 { font-size: 18pt; font-weight: bold; margin-top: 12pt; margin-bottom: 6pt; color: #1f4e79; }
          h2 { font-size: 14pt; font-weight: bold; margin-top: 10pt; margin-bottom: 4pt; color: #2e75b6; }
          h3 { font-size: 12pt; font-weight: bold; margin-top: 8pt; margin-bottom: 2pt; color: #5b9bd5; }
          p { margin: 0 0 6pt 0; line-height: inherit; }
          .text-align-center, [align="center"] { text-align: center !important; }
          .text-align-right, [align="right"] { text-align: right !important; }
          .text-align-justify, [align="justify"] { text-align: justify !important; }
          .text-align-left, [align="left"] { text-align: left !important; }
          table { border-collapse: collapse; width: 100%; margin: 8pt 0; border: 1px solid #bfbfbf; }
          th, td { border: 1px solid #bfbfbf; padding: 6pt 10pt; vertical-align: top; }
          th { background-color: #f2f2f2; font-weight: bold; }
          img { max-width: 100%; height: auto; }
          .page-break { page-break-after: always; mso-break-type: section-break; }
        </style>
      </head>
      <body>
        ${cleanHtml}
      </body>
      </html>
    `;

    try {
      const rawData = await asBlob(fullHtml, {
        orientation: options.orientation || 'portrait',
        margins: options.margins || { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      });

      let finalBlob;
      if (rawData instanceof Blob) {
        finalBlob = rawData;
      } else if (rawData instanceof ArrayBuffer) {
        finalBlob = new Blob([rawData], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      } else if (rawData && rawData.buffer instanceof ArrayBuffer) {
        finalBlob = new Blob([rawData.buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      } else {
        finalBlob = new Blob([new Uint8Array(rawData)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      }

      // Clean filename
      let cleanBaseName = fileName.replace(/\.docx$/i, '').trim();
      if (!cleanBaseName) cleanBaseName = 'TaiLieu';
      const finalName = `${cleanBaseName}.docx`;

      // Method 1: Modern File System Access API (Native OS Save Dialog with accurate .docx extension)
      if (typeof window.showSaveFilePicker === 'function') {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: finalName,
            types: [
              {
                description: 'Tài liệu Microsoft Word (*.docx)',
                accept: {
                  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx']
                }
              }
            ]
          });
          const writable = await handle.createWritable();
          await writable.write(finalBlob);
          await writable.close();
          return { success: true, fileName: finalName };
        } catch (pickerErr) {
          if (pickerErr.name === 'AbortError') {
            return { cancelled: true };
          }
          console.warn('showSaveFilePicker error, falling back to download:', pickerErr);
        }
      }

      // Method 2: FileSaver.js saveAs with a named File object
      const finalFile = new File([finalBlob], finalName, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      try {
        saveAs(finalFile, finalName);
      } catch (saveAsErr) {
        console.warn('FileSaver failed, falling back to persistent anchor:', saveAsErr);
        const url = URL.createObjectURL(finalBlob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.setAttribute('download', finalName);
        a.download = finalName;
        document.body.appendChild(a);
        a.click();
        // Do NOT remove anchor immediately so Chrome does not lose download metadata
        setTimeout(() => {
          if (a.parentNode) a.parentNode.removeChild(a);
          URL.revokeObjectURL(url);
        }, 60000);
      }

      return { success: true, fileName: finalName };
    } catch (err) {
      console.error('Lỗi khi xuất DOCX:', err);
      throw err;
    }
  }

  /**
   * Save draft to LocalStorage
   */
  static saveToLocalStorage(htmlContent, title) {
    try {
      const payload = {
        title: title || 'Tài liệu 1',
        content: htmlContent,
        updatedAt: new Date().toISOString()
      };
      localStorage.setItem('word_editor_draft', JSON.stringify(payload));
      return true;
    } catch (e) {
      console.warn('LocalStorage save failed:', e);
      return false;
    }
  }

  /**
   * Load draft from LocalStorage
   */
  static loadFromLocalStorage() {
    try {
      const data = localStorage.getItem('word_editor_draft');
      if (data) {
        return JSON.parse(data);
      }
    } catch (e) {
      console.warn('LocalStorage load failed:', e);
    }
    return null;
  }
}
