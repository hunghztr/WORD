import {
  DecoupledEditor,
  Plugin,
  GeneralHtmlSupport,
  Essentials,
  Paragraph,
  Heading,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Subscript,
  Superscript,
  FontFamily,
  FontSize,
  FontColor,
  FontBackgroundColor,
  Alignment,
  List,
  TodoList,
  Indent,
  IndentBlock,
  Table,
  TableToolbar,
  TableProperties,
  TableCellProperties,
  TableCaption,
  HorizontalLine,
  PageBreak,
  Link,
  Image,
  ImageToolbar,
  ImageCaption,
  ImageStyle,
  ImageResize,
  ImageUpload,
  Base64UploadAdapter,
  WordCount,
  Undo
} from 'ckeditor5';
import 'ckeditor5/ckeditor5.css';

/**
 * Custom LineHeight Plugin for CKEditor 5 Decoupled
 * Handles line-height and data-lineheight upcast & downcast
 */
class LineHeightEditing extends Plugin {
  static get pluginName() {
    return 'LineHeightEditing';
  }

  init() {
    const editor = this.editor;
    const schema = editor.model.schema;
    const conversion = editor.conversion;

    schema.extend('$block', {
      allowAttributes: ['lineHeight']
    });
    schema.setAttributeProperties('lineHeight', {
      isFormatting: true
    });

    // Upcast: read line-height from data-lineheight (preferred) or inline style on blocks.
    // attributeToAttribute consumes the values before GeneralHtmlSupport can keep a raw
    // copy of them (which would bypass this plugin and override its rendering).
    const readLineHeight = viewElement => {
      const lh = (viewElement.getAttribute('data-lineheight') || viewElement.getStyle('line-height') || '').trim();
      return lh && lh !== 'normal' ? lh : null;
    };
    conversion.for('upcast').attributeToAttribute({
      view: { key: 'data-lineheight' },
      model: { key: 'lineHeight', value: readLineHeight }
    });
    conversion.for('upcast').attributeToAttribute({
      view: { styles: { 'line-height': /[\s\S]+/ } },
      model: { key: 'lineHeight', value: readLineHeight }
    });

    // Downcast: model lineHeight -> view style="line-height: ..." + data-lineheight.
    // Word's "multiple" spacing is relative to the font's natural line height, while CSS
    // unitless line-height is relative to font-size. The editing view scales unitless
    // values by --word-lh-factor (set per document font) so lines look like in Word;
    // the data output keeps the plain Word value.
    const addLineHeightDowncast = (group, toCss) => {
      conversion.for(group).add(dispatcher => {
        dispatcher.on('attribute:lineHeight', (evt, data, conversionApi) => {
          if (!conversionApi.consumable.consume(data.item, evt.name)) return;
          const viewWriter = conversionApi.writer;
          const viewElement = conversionApi.mapper.toViewElement(data.item);
          if (!viewElement) return;

          if (data.attributeNewValue) {
            viewWriter.setStyle('line-height', toCss(data.attributeNewValue), viewElement);
            viewWriter.setAttribute('data-lineheight', data.attributeNewValue, viewElement);
          } else {
            viewWriter.removeStyle('line-height', viewElement);
            viewWriter.removeAttribute('data-lineheight', viewElement);
          }
        }, { priority: 'low' });
      });
    };
    addLineHeightDowncast('dataDowncast', value => value);
    addLineHeightDowncast('editingDowncast', value =>
      /^\d*\.?\d+$/.test(value) ? `calc(${value} * var(--word-lh-factor, 1))` : value
    );
  }
}

/**
 * Measure a font's natural line height ("single" spacing in Word) as a multiple of its size.
 * @param {string} fontFamily
 * @returns {number}
 */
export function measureLineHeightFactor(fontFamily) {
  // A block's height with line-height:normal = ascent + descent + line gap, like Word
  const probe = document.createElement('div');
  probe.style.cssText = `position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;line-height:normal;font-family:${fontFamily};`;
  probe.textContent = 'ÂgHịỹ';
  document.body.appendChild(probe);
  const factor = probe.getBoundingClientRect().height / 100;
  probe.remove();
  return factor > 0.9 && factor < 1.6 ? Math.round(factor * 1000) / 1000 : 1.15;
}



/**
 * Custom AlignmentStyling Plugin
 * Ensures paragraph text-align styles, align attributes, and text-align-* classes
 * are seamlessly two-way converted and preserved.
 */
class AlignmentStyling extends Plugin {
  static get pluginName() {
    return 'AlignmentStyling';
  }

  init() {
    const editor = this.editor;
    const conversion = editor.conversion;

    // Upcast all variations of alignment to model 'alignment'
    for (const align of ['left', 'center', 'right', 'justify']) {
      conversion.for('upcast').attributeToAttribute({
        view: {
          styles: {
            'text-align': align
          }
        },
        model: {
          key: 'alignment',
          value: align
        }
      });
      conversion.for('upcast').attributeToAttribute({
        view: {
          key: 'class',
          value: `text-align-${align}`
        },
        model: {
          key: 'alignment',
          value: align
        }
      });
      conversion.for('upcast').attributeToAttribute({
        view: {
          key: 'align',
          value: align
        },
        model: {
          key: 'alignment',
          value: align
        }
      });
    }

    // Downcast model alignment to view style, class, and align attribute
    conversion.for('downcast').add(dispatcher => {
      dispatcher.on('attribute:alignment', (evt, data, conversionApi) => {
        if (!conversionApi.consumable.consume(data.item, evt.name)) return;
        const viewWriter = conversionApi.writer;
        const viewElement = conversionApi.mapper.toViewElement(data.item);
        if (!viewElement) return;

        // Clean previous alignment classes
        for (const al of ['left', 'center', 'right', 'justify']) {
          viewWriter.removeClass(`text-align-${al}`, viewElement);
        }

        if (data.attributeNewValue) {
          viewWriter.setStyle('text-align', data.attributeNewValue, viewElement);
          viewWriter.addClass(`text-align-${data.attributeNewValue}`, viewElement);
          viewWriter.setAttribute('align', data.attributeNewValue, viewElement);
        } else {
          viewWriter.removeStyle('text-align', viewElement);
          viewWriter.removeAttribute('align', viewElement);
        }
      }, { priority: 'low' });
    });
  }
}

/**
 * Initialize CKEditor 5 Decoupled Editor with full Word-like plugins
 * @param {HTMLElement} editorElement 
 * @param {HTMLElement} toolbarContainer 
 * @param {Function} onWordCountUpdate 
 * @returns {Promise<DecoupledEditor>}
 */
export async function initWordEditor(editorElement, toolbarContainer, onWordCountUpdate) {
  const editor = await DecoupledEditor.create(editorElement, {
    licenseKey: 'GPL',
    plugins: [
      GeneralHtmlSupport,
      LineHeightEditing,
      AlignmentStyling,
      Essentials,
      Paragraph,
      Heading,
      Bold,
      Italic,
      Underline,
      Strikethrough,
      Subscript,
      Superscript,
      FontFamily,
      FontSize,
      FontColor,
      FontBackgroundColor,
      Alignment,
      List,
      TodoList,
      Indent,
      IndentBlock,
      Table,
      TableToolbar,
      TableProperties,
      TableCellProperties,
      TableCaption,
      HorizontalLine,
      PageBreak,
      Link,
      Image,
      ImageToolbar,
      ImageCaption,
      ImageStyle,
      ImageResize,
      ImageUpload,
      Base64UploadAdapter,
      WordCount,
      Undo
    ],
    htmlSupport: {
      allow: [
        {
          name: /.*/,
          attributes: true,
          classes: true,
          styles: true
        }
      ]
    },
    toolbar: {
      items: [
        'undo', 'redo',
        '|',
        'heading',
        '|',
        'fontFamily', 'fontSize',
        '|',
        'bold', 'italic', 'underline', 'strikethrough', 'subscript', 'superscript',
        '|',
        'fontColor', 'fontBackgroundColor',
        '|',
        'alignment:left', 'alignment:center', 'alignment:right', 'alignment:justify',
        '|',
        'bulletedList', 'numberedList', 'todoList',
        'outdent', 'indent',
        '|',
        'link', 'insertImage', 'insertTable', 'horizontalLine', 'pageBreak'
      ],
      shouldNotGroupWhenFull: true
    },
    heading: {
      options: [
        { model: 'paragraph', title: 'Văn bản thường (Normal)', class: 'ck-heading_paragraph' },
        { model: 'heading1', view: 'h1', title: 'Tiêu đề 1 (Heading 1)', class: 'ck-heading_heading1' },
        { model: 'heading2', view: 'h2', title: 'Tiêu đề 2 (Heading 2)', class: 'ck-heading_heading2' },
        { model: 'heading3', view: 'h3', title: 'Tiêu đề 3 (Heading 3)', class: 'ck-heading_heading3' }
      ]
    },
    fontFamily: {
      options: [
        'default',
        'Calibri, sans-serif',
        'Times New Roman, Times, serif',
        'Arial, Helvetica, sans-serif',
        'Segoe UI, Tahoma, sans-serif',
        'Georgia, serif',
        'Garamond, serif',
        'Courier New, Courier, monospace'
      ],
      supportAllValues: true
    },
    fontSize: {
      options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72],
      supportAllValues: true
    },
    fontColor: {
      columns: 5,
      documentColors: 10
    },
    fontBackgroundColor: {
      columns: 5,
      documentColors: 10
    },
    alignment: {
      options: [
        { name: 'left', className: 'text-align-left' },
        { name: 'center', className: 'text-align-center' },
        { name: 'right', className: 'text-align-right' },
        { name: 'justify', className: 'text-align-justify' }
      ]
    },
    table: {
      contentToolbar: [
        'tableColumn', 'tableRow', 'mergeTableCells',
        'tableProperties', 'tableCellProperties'
      ]
    },
    image: {
      toolbar: [
        'imageStyle:inline', 'imageStyle:block', 'imageStyle:side',
        '|',
        'toggleImageCaption', 'imageTextAlternative',
        '|',
        'imageResize'
      ]
    }
  });

  // Attach the CKEditor toolbar to the container
  if (toolbarContainer) {
    toolbarContainer.appendChild(editor.ui.view.toolbar.element);
  }

  // Hook word count plugin if callback provided
  if (onWordCountUpdate) {
    const wordCountPlugin = editor.plugins.get('WordCount');
    wordCountPlugin.on('change:words', (evt, propertyName, newValue) => {
      onWordCountUpdate({
        words: newValue,
        characters: wordCountPlugin.characters
      });
    });
    // Initial trigger
    onWordCountUpdate({
      words: wordCountPlugin.words,
      characters: wordCountPlugin.characters
    });
  }

  // Dynamic labels for Font Size & Font Family in Word Ribbon style
  const setupFontDropdownLabels = () => {
    const fontSizeCmd = editor.commands.get('fontSize');
    const fontFamilyCmd = editor.commands.get('fontFamily');

    const updateFontSizeText = () => {
      const dropdownEl = editor.ui.view.toolbar.element?.querySelector('.ck-font-size-dropdown');
      if (dropdownEl) {
        let labelEl = dropdownEl.querySelector('.ck-button__label');
        if (!labelEl) {
          const btn = dropdownEl.querySelector('.ck-dropdown__button');
          if (btn) {
            labelEl = document.createElement('span');
            labelEl.className = 'ck-button__label';
            const arrow = btn.querySelector('.ck-dropdown__arrow');
            if (arrow) {
              btn.insertBefore(labelEl, arrow);
            } else {
              btn.appendChild(labelEl);
            }
          }
        }
        if (labelEl) {
          const currentVal = fontSizeCmd?.value;
          labelEl.textContent = currentVal ? String(currentVal).replace(/px|pt/i, '') : '14';
        }
      }
    };

    const updateFontFamilyText = () => {
      const dropdownEl = editor.ui.view.toolbar.element?.querySelector('.ck-font-family-dropdown');
      if (dropdownEl) {
        let labelEl = dropdownEl.querySelector('.ck-button__label');
        if (!labelEl) {
          const btn = dropdownEl.querySelector('.ck-dropdown__button');
          if (btn) {
            labelEl = document.createElement('span');
            labelEl.className = 'ck-button__label';
            const arrow = btn.querySelector('.ck-dropdown__arrow');
            if (arrow) {
              btn.insertBefore(labelEl, arrow);
            } else {
              btn.appendChild(labelEl);
            }
          }
        }
        if (labelEl) {
          const currentVal = fontFamilyCmd?.value;
          const familyName = currentVal ? currentVal.split(',')[0].replace(/['"]/g, '').trim() : 'Calibri';
          labelEl.textContent = familyName;
        }
      }
    };

    fontSizeCmd?.on('change:value', updateFontSizeText);
    fontFamilyCmd?.on('change:value', updateFontFamilyText);

    // Initial update after DOM attach
    setTimeout(() => {
      updateFontSizeText();
      updateFontFamilyText();
    }, 150);
  };

  setupFontDropdownLabels();

  return editor;
}
