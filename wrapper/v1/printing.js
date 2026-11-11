(function (global) {
  'use strict';

  const PRINT_DPI = 300;
  let activePrintJob = null;

  function createPage(win, width, height, fontManager) {
    const hasValidSize = Number.isFinite(width) && Number.isFinite(height) &&
      width > 0 && height > 0;

    if (!hasValidSize) {
      throw new Error('Invalid print page size');
    }

    const canvas = win.document.createElement('canvas');
    canvas.width = Math.ceil(width * PRINT_DPI / 25.4);
    canvas.height = Math.ceil(height * PRINT_DPI / 25.4);

    const context = canvas.getContext('2d');
    context.fillStyle = '#fff';
    context.fillRect(0, 0, canvas.width, canvas.height);

    const graphics = new win.AscCommon.CGraphics();
    graphics.init(context, canvas.width, canvas.height, width, height);
    graphics.m_oFontManager = fontManager;
    graphics.isPrintMode = true;
    graphics.IsPrintMode = true;
    graphics.IsNoDrawingEmptyPlaceholder = true;
    graphics.transform(1, 0, 0, 1, 0, 0);
    graphics.BeginPage = function () {};
    graphics.EndPage = function () {};

    return { canvas, graphics, width, height };
  }

  function withSpreadsheetScale(win, workbook, draw) {
    const browser = win.AscCommon.AscBrowser;
    const ratio = browser.retinaPixelRatio;
    const zoom = workbook.getZoom();

    try {
      browser.retinaPixelRatio = 1;
      workbook.changeZoom(1, true, true);
      workbook.changeZoom(null, true, true);

      return draw();
    } finally {
      browser.retinaPixelRatio = ratio;
      workbook.changeZoom(null, true, true);
      workbook.changeZoom(zoom, true, true);
    }
  }

  function createSpreadsheetPrintRenderer(graphics) {
    // A raw CGraphics selects the SDK's screen-preview path, where clipping
    // uses screen pixels. Keep its PDF print path in millimetres at any DPI.
    return new Proxy({}, {
      get: function (_, property) {
        if (property === 'SetBaseTransform') {
          return function (transform) {
            if (transform) {
              graphics.SetBaseTransform(transform);
            } else {
              graphics.ResetBaseTransform();
            }
          };
        }

        const value = graphics[property];

        return typeof value === 'function' ? value.bind(graphics) : value;
      },
      set: function (_, property, value) {
        graphics[property] = value;

        return true;
      }
    });
  }

  function renderSpreadsheetPage(win, api, page, index, pages) {
    return withSpreadsheetScale(win, api.wb, function () {
      const defaults = win.AscCommon.c_oAscPrintDefaultSettings;
      const result = createPage(win,
        page ? page.pageWidth : defaults.PageWidth,
        page ? page.pageHeight : defaults.PageHeight,
        api.wb.fmgrGraphics[3]);
      const printer = new win.AscCommonExcel.CPdfPrinter(
        api.wb.fmgrGraphics[3], api.wb.m_oFont.clone());
      printer.DocumentRenderer = createSpreadsheetPrintRenderer(result.graphics);

      // The SDK's PDF printer writes image commands to a binary stream.
      // Draw them on our canvas while keeping its normal text/layout code.
      printer.drawImage = function (source, sx, sy, sw, sh, dx, dy, dw, dh, iw, ih) {
        if (this.bIsSimpleCommands) {
          result.graphics.drawImage(source, sx, sy, sw, sh, dx, dy);
        } else {
          const crop = {
            l: sx / iw * 100, t: sy / ih * 100,
            r: (sx + sw) / iw * 100, b: (sy + sh) / ih * 100
          };
          result.graphics.drawImage(source,
            dx * 25.4 / 96, dy * 25.4 / 96,
            dw * 25.4 / 96, dh * 25.4 / 96, 255, crop);
        }

        return this;
      };

      const worksheet = page ? api.wb.getWorksheet(page.indexWorksheet) : api.wb.getWorksheet();
      worksheet.drawForPrint(printer, page, index, pages);

      return result;
    });
  }

  function renderDocumentPage(win, api, type, index) {
    const model = api.WordControl.m_oLogicDocument;
    const page = type === 'slide'
      ? { width_mm: model.GetWidthMM(), height_mm: model.GetHeightMM() }
      : api.WordControl.m_oDrawingDocument.m_arrPages[index];
    const result = createPage(win, page.width_mm, page.height_mm,
      win.AscCommon.g_fontManager);
    const showMarks = api.ShowParaMarks;

    try {
      api.ShowParaMarks = false;

      if (type === 'word') {
        model.SetupBeforeNativePrint({
          isPrint: true, drawPlaceHolders: false, drawFormHighlight: false
        }, result.graphics);
      }

      model.DrawPage(index, result.graphics);
    } finally {
      api.ShowParaMarks = showMarks;

      if (type === 'word') {
        model.RestoreAfterNativePrint();
      }
    }

    return result;
  }

  function checkAllowed(options) {
    if (options.isAllowed && !options.isAllowed()) {
      const error = new Error('Printing is no longer available');
      error.name = 'AbortError';

      throw error;
    }
  }

  async function prepare(win, type, options = {}) {
    checkAllowed(options);

    const api = win.Asc && win.Asc.editor;

    if (!api || !api.isDocumentLoadComplete) {
      throw new Error('Document is not ready for printing');
    }

    if (type === 'cell' && !api.asc_closeCellEditor()) {
      throw new Error('Finish editing the cell before printing');
    }

    const frame = global.document.createElement('iframe');
    frame.id = 'sk-print-frame';
    frame.setAttribute('aria-hidden', 'true');
    frame.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;' +
      'border:0;visibility:hidden;';
    const urls = [];
    const cleanup = function () {
      urls.forEach(function (url) { global.URL.revokeObjectURL(url); });
      frame.remove();
    };
    global.document.body.appendChild(frame);

    try {
      const doc = frame.contentDocument;
      doc.title = options.title || 'Document';
      const style = doc.createElement('style');
      style.textContent = 'html,body{margin:0;padding:0;background:white;}' +
        '.print-page{break-after:page;overflow:hidden;}' +
        '.print-page:last-child{break-after:auto;}' +
        '.print-page img{display:block;width:100%;height:100%;}' +
        '@page{margin:0;}';
      doc.head.appendChild(style);

      let pages;

      if (type === 'cell') {
        const adjust = new win.Asc.asc_CAdjustPrint();
        adjust.asc_setPrintType(win.Asc.c_oAscPrintType.ActiveSheets);
        pages = withSpreadsheetScale(win, api.wb, function () {
          return api.wb.calcPagesPrint(adjust).arrPages;
        });

        if (!pages.length) {
          pages = [null];
        }
      } else {
        const count = type === 'slide'
          ? api.WordControl.m_oLogicDocument.Slides.length
          : api.WordControl.m_oDrawingDocument.m_lPagesCount;
        pages = Array.from({ length: count }, function (_, index) { return index; });
      }

      if (!pages.length) {
        throw new Error('There are no pages to print');
      }

      for (let index = 0; index < pages.length; index++) {
        checkAllowed(options);

        const result = type === 'cell'
          ? renderSpreadsheetPage(win, api, pages[index], index, pages)
          : renderDocumentPage(win, api, type, index);
        const blob = await new Promise(function (resolve, reject) {
          result.canvas.toBlob(function (image) {
            if (image) {
              resolve(image);
            } else {
              reject(new Error('Print page could not be rendered'));
            }
          }, 'image/png');
        });
        result.canvas.width = 0;
        result.canvas.height = 0;

        const url = global.URL.createObjectURL(blob);
        urls.push(url);
        const pageName = 'sk-page-' + index;
        style.textContent += '@page ' + pageName + '{size:' + result.width +
          'mm ' + result.height + 'mm;margin:0;}';
        const section = doc.createElement('section');
        section.className = 'print-page';
        section.style.cssText = 'page:' + pageName + ';width:' + result.width +
          'mm;height:' + result.height + 'mm;';
        const image = doc.createElement('img');
        image.alt = '';
        image.src = url;
        section.appendChild(image);
        doc.body.appendChild(section);
        await image.decode();
      }

      checkAllowed(options);

      return { frame, cleanup, pageCount: pages.length };
    } catch (error) {
      cleanup();

      throw error;
    }
  }

  async function print(win, type, options) {
    if (activePrintJob) {
      activePrintJob.cleanup();
      activePrintJob = null;
    }

    const job = await prepare(win, type, options);
    activePrintJob = job;
    job.frame.contentWindow.addEventListener('afterprint', function () {
      job.cleanup();

      if (activePrintJob === job) {
        activePrintJob = null;
      }
    }, { once: true });

    try {
      job.frame.contentWindow.focus();
      job.frame.contentWindow.print();
    } catch (error) {
      job.cleanup();
      activePrintJob = null;

      throw error;
    }
  }

  global.OfficePrinting = { prepare, print };
})(window);
