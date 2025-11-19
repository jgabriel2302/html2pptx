'use strict';
/**
 * @file HTML2PPTX.js
 * @summary Helper utilities for exporting the LATAM Padronização slides to PowerPoint.
 * @description Single-class implementation that reads rendered SVG/table content and recreates it via pptxgen.js, ready for use in browsers, CommonJS, AMD or ES-module environments.
 *
 * @author João Gabriel Corrêa da Silva (https://github.com/jgabriel2302/)
 * @version 1.0.0
 * @license MIT
 */
/**
 * @summary Utility for converting rendered HTML/SVG slides into PPTX decks.
 * @description Works as either an ES module (`import HTML2PPTX from './HTML2PPTX.js';`) or a classic `<script>` that exposes a global constructor.
 * @class HTML2PPTX
 * @module HTML2PPTX
 */
class HTML2PPTX {
  /**
   * @summary Minimum stroke width before ignoring line data to avoid artifacts.
   * @constant
   * @returns {number} Stroke width threshold measured in EMU.
   */
  static get STROKE_LIMIT() {
    return 9525;
  }

  /**
   * @summary Set of SVG tag names treated as vector primitives during export.
   * @constant
   * @returns {Set<string>} Cached set of uppercase tag names.
   */
  static get SVG_ELEMENTS() {
    if (!this.#SVG_ELEMENTS) {
      this.#SVG_ELEMENTS = new Set(['TEXT', 'RECT', 'G']);
    }
    return this.#SVG_ELEMENTS;
  }

  /**
   * @summary Mapping of DOM `text-align`/`text-anchor` values to pptxgen alignment options.
   * @constant
   * @returns {Record<'start'|'end'|'left'|'right'|'center'|'middle'|'justify', {align: string, autoFit: boolean}>} Alignment dictionary reused across slides.
   */
  static get TEXT_ALIGNS() {
    if (!this.#TEXT_ALIGNS) {
      this.#TEXT_ALIGNS = {
        start: { align: 'left', autoFit: false },
        end: { align: 'right', autoFit: false },
        left: { align: 'left', autoFit: false },
        right: { align: 'right', autoFit: false },
        center: { align: 'center', autoFit: false },
        middle: { align: 'center', autoFit: false },
        justify: { align: 'left', autoFit: true },
      };
    }
    return this.#TEXT_ALIGNS;
  }

  static #SVG_ELEMENTS;
  /**
   * @constant 
   * @type {Record<'start'|'end'|'left'|'right'|'center'|'middle'|'justify', {align: string, autoFit: boolean}>} Alignment dictionary reused across slides.
   */
  static #TEXT_ALIGNS;

  #pptx = null;
  /**
   * @summary Creates an exporter instance with metadata, layout definition and optional editor context.
   * @param {Object} [options={}] Configuration object.
   * @param {string} [options.author=''] Author metadata passed to pptxgen.
   * @param {string} [options.company=''] Company metadata passed to pptxgen.
   * @param {string} [options.title=''] Title metadata passed to pptxgen.
   * @param {string} [options.subject=''] Subject metadata passed to pptxgen.
   * @param {string} [options.fileName] Custom filename for the exported PPTX.
   * @param {string} [options.locale='pt-BR'] Locale used for date/time formatting.
   * @param {HTMLElement} [options.editor] Container used to detect elements hidden only during presentation preview.
   * @param {{name?: string, width?: number, height?: number}} [options.layout] PPTX layout definition (in inches).
   * @param {'fit'|'viewport'} [options.slideSizing='fit'] Defines whether the SVG is scaled to fit the slide (`'fit'`) or if its viewport is treated as the slide bounds (`'viewport'`).
   * @property {number} slideWidthEmu Width of the slide in EMUs derived from the layout.
   * @property {number} slideHeightEmu Height of the slide in EMUs derived from the layout.
   */
  constructor(options = {}) {
    this.author =
      options.author ?? '';
    this.company = options.company ?? '';
    this.title = options.title ?? '';
    this.subject = options.subject ?? '';
    this.fileName = options.fileName ?? options.title + '-presentation.pptx';
    this.locale = options.locale ?? 'pt-BR';
    this.editor =
      options.editor ??
      (typeof document !== 'undefined'
        ? document.getElementById('result')
        : null);
   this.layout = {
      name: options.layout?.name ?? 'HTMLTOPPTX-16x9',
      width: options.layout?.width ?? 20,
      height: options.layout?.height ?? 11.25,
    };
   this.slideWidthEmu = this.layout.width * HTML2PPTX.EMU_PER_IN;
   this.slideHeightEmu = this.layout.height * HTML2PPTX.EMU_PER_IN;
   this.slideSizing = this.#normalizeSlideSizing(options.slideSizing);
  }
  
  /**
   * @summary Converts DOM slides into PPTX slides.
   * @description Iterates over SVG nodes, converting each to pptxgen shapes/text. Can append to an existing presentation.
   * @param {Iterable<SVGElement>|SVGElement|null} slidesSvg Collection of SVG elements or a single SVG node.
   * @returns {HTML2PPTX} The HTML2PPTX instance, useful when chaining additional operations.
   */
  generate(slidesSvg) {
    this.#pptx = this.#pptx ?? this.#createPresentation();
    const nodes = this.#normalizeSlides(slidesSvg);
    for (const svg of nodes) {
      this.#renderSlide(this.#pptx, svg);
    }
    return this;
  }

  /**
   * @summary Writes the pptx file.
   * @description Writes the pptx file from the pptx in context using the fileName setted in the options.
   * @returns {HTML2PPTX} The HTML2PPTX instance, useful when chaining additional operations.
   */
  download(){
    if(!this.#pptx) return this;
    this.#pptx.writeFile({ fileName: this.fileName });
    return this;
  }

  /**
   * @summary Creates a pptxgen presentation configured with metadata and layout.
   * @private
   * @inner
   * @returns {PptxGenJS} Fresh pptxgen presentation instance.
   */
  #createPresentation() {
    const pptx = new PptxGenJS();
    pptx.author = this.author;
    pptx.company = this.company;
    pptx.title = this.title;
    pptx.subject = this.subject;
    pptx.revision = String(Date.now());
    pptx.defineLayout(this.layout);
    pptx.layout = this.layout.name;
    return pptx;
  }

  /**
   * @summary Normalizes any iterable input into an array of SVG elements.
   * @private
   * @inner
   * @param {Iterable<SVGElement>|SVGElement|null|undefined} slidesSvg Arbitrary iterable or single node reference.
   * @returns {SVGElement[]} Array of SVG nodes ready for export.
   */
  #normalizeSlides(slidesSvg) {
    if (!slidesSvg) return [];
    if (slidesSvg instanceof NodeList || Array.isArray(slidesSvg)) {
      return Array.from(slidesSvg);
    }
    if (
      typeof slidesSvg[Symbol.iterator] === 'function' &&
      !slidesSvg.tagName
    ) {
      return Array.from(slidesSvg);
    }
    return slidesSvg.tagName ? [slidesSvg] : [];
  }

  /**
   * @summary Resolves the most appropriate bounding rectangle for a DOM element.
   * @private
   * @inner
   * @param {Element} element Target element.
   * @param {{slideSizing: 'fit'|'viewport', svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}, svg: SVGElement}} context Rendering context.
   * @returns {{x:number,y:number,width:number,height:number,coordinateSpace:'svg'|'screen'}} Rectangle in SVG or screen coordinates.
   */
  #resolveElementRect(element, context) {
    if (context.slideSizing === 'viewport') {
      const svgRect = this.#getSvgRect(element, context);
      if (svgRect) return svgRect;
    }
    const clientRect = this.#getClientRect(element);
    if (clientRect) return clientRect;
    const rect = element.getBoundingClientRect();
    return {
      x: rect.x ?? rect.left ?? 0,
      y: rect.y ?? rect.top ?? 0,
      width:
        rect.width ??
        Math.max(
          0,
          (rect.right ?? rect.left ?? 0) - (rect.left ?? rect.x ?? 0)
        ),
      height:
        rect.height ??
        Math.max(
          0,
          (rect.bottom ?? rect.top ?? 0) - (rect.top ?? rect.y ?? 0)
        ),
      coordinateSpace: 'screen',
    };
  }

  /**
   * @summary Normalizes the slide sizing mode coming from options or dataset attributes.
   * @private
   * @inner
   * @param {string} [mode='fit'] Desired mode.
   * @default 'fit'
   * @returns {'fit'|'viewport'} Sanitized mode string.
   */
  #normalizeSlideSizing(mode) {
    if (typeof mode === 'string') {
      const normalized = mode.trim().toLowerCase();
      if (normalized === 'viewport' || normalized === 'fit') {
        return normalized;
      }
    }
    return 'fit';
  }

  /**
   * @summary Attempts to compute the bounding box of an element in the parent SVG coordinate space.
   * @private
   * @inner
   * @param {Element} element Target SVG/HTML element.
   * @param {{svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}, svg: SVGElement}} context Slide context.
   * @returns {{x:number,y:number,width:number,height:number,coordinateSpace:'svg'}|null} Bounding rectangle or null when unavailable.
   */
  #getSvgRect(element, context) {
    if (
      !(element instanceof SVGElement) ||
      typeof element.getBBox !== 'function' ||
      typeof element.getScreenCTM !== 'function' ||
      typeof DOMPoint !== 'function'
    ) {
      return null;
    }
    try {
      const bbox = element.getBBox();
      const matrix = element.getScreenCTM();
      if (!bbox || !matrix) return null;

      const corners = [
        new DOMPoint(bbox.x, bbox.y),
        new DOMPoint(bbox.x + bbox.width, bbox.y),
        new DOMPoint(bbox.x, bbox.y + bbox.height),
        new DOMPoint(bbox.x + bbox.width, bbox.y + bbox.height),
      ].map((point) => point.matrixTransform(matrix));
      const xs = corners.map((point) => point.x);
      const ys = corners.map((point) => point.y);

      const minDomX = Math.min(...xs);
      const minDomY = Math.min(...ys);
      const maxDomX = Math.max(...xs);
      const maxDomY = Math.max(...ys);

      const { svgRect, viewBox } = context;
      const domWidth = svgRect.width || 1;
      const domHeight = svgRect.height || 1;
      const domToViewX = (value) =>
        (value / domWidth) * (viewBox.width || domWidth);
      const domToViewY = (value) =>
        (value / domHeight) * (viewBox.height || domHeight);

      const viewX = domToViewX(minDomX - (svgRect.x ?? svgRect.left ?? 0)) + (viewBox.minX ?? 0);
      const viewY = domToViewY(minDomY - (svgRect.y ?? svgRect.top ?? 0)) + (viewBox.minY ?? 0);
      const viewW = domToViewX(maxDomX - minDomX);
      const viewH = domToViewY(maxDomY - minDomY);

      return {
        x: viewX,
        y: viewY,
        width: viewW,
        height: viewH,
        coordinateSpace: 'svg',
      };
    } catch (error) {
      return null;
    }
  }

  /**
   * @summary Creates a PPTX slide and processes relevant SVG children preserving the aspect ratio.
   * @private
   * @inner
   * @param {PptxGenJS} pptx Target pptxgen presentation.
   * @param {SVGElement} svg Source SVG element representing a slide. Supports a `data-slide-sizing="viewport|fit"` attribute to override the global behavior.
   * @returns {void}
   */
  #renderSlide(pptx, svg) {
    const slide = pptx.addSlide();
    const svgRect = svg.getBoundingClientRect();
    const viewBox = this.#getViewBox(svg, svgRect);
    const slideSizing = this.#normalizeSlideSizing(
      svg.dataset?.slideSizing ?? this.slideSizing
    );
    const context = { svgRect, viewBox, slide, slideSizing, svg };
    const elements = svg.querySelectorAll(
      'text,rect,circle,line,path,polygon,g[name],div,li,td'
    );
    for (const element of elements) {
      this.#renderElement(element, context, pptx);
    }
  }

  /**
   * @summary Converts a single DOM element into pptxgen shapes, text or images.
   * @private
   * @inner
   * @param {Element} element DOM element being exported.
   * @param {{svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}, slideSizing: 'fit'|'viewport', slide: PptxGenJS.Slide, svg: SVGElement}} context Precomputed slide context.
   * @param {PptxGenJS} pptx 
   * @returns {void}
   */
  #renderElement(element, context, pptx) {
    if (this.#shouldSkipElement(element)) return;

    const rect = this.#resolveElementRect(element, context);
    const metrics = this.#normalizeMetrics(
      this.#rectToSlideMetrics(rect, context),
      context
    );
    if (metrics.w <= 0 || metrics.h <= 0) return;
    const style = window.getComputedStyle(element);
    if (HTML2PPTX.isHidden(element)) return;

    const colors = this.#resolveColors(element, style);
    const borderWidth = this.#resolveBorderWidth(style);
    const font = this.#resolveFont(style);
    const alignment =
      HTML2PPTX.TEXT_ALIGNS[style.getPropertyValue('text-align')] ??
      HTML2PPTX.TEXT_ALIGNS[style.getPropertyValue('text-anchor')] ??
      HTML2PPTX.TEXT_ALIGNS.center;
    const dashType = style.getPropertyValue('stroke-dasharray')
      ? 'dashDot'
      : 'solid';
    const radius = this.#resolveRadius(element, style, rect);
    const textMetrics = this.#textBoxMetrics(metrics, style);

    const tag = element.tagName.toUpperCase();
    const name = element.getAttribute('name');
    const slide = context.slide;

    if (name === '@updateDate') {
      const dateBounds = {
        x: textMetrics.x,
        y: textMetrics.y,
        h: textMetrics.h,
        w: textMetrics.w * 2,
      };
      const dateOptions = {
        ...this.#applyPercentPosition(dateBounds),
        h: dateBounds.h,
        w: dateBounds.w,
        breakLine: false,
        color: colors.fill.hex,
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        bold: font.bold,
        ...alignment,
      };
      slide.addText(
        this.#formatDate(),
        this.#applyPercentPosition({
          ...dateOptions,
          x: this.#alignX(dateOptions, dateBounds),
        })
      );
      return;
    }

    if (name === '@pageNumber') {
      slide.slideNumber = {
        ...this.#applyPercentPosition(textMetrics),
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        color: colors.fill.hex,
      };
      return;
    }

    if (name === '@Logo' || name === '@Svg') {
      const temp = document.createElement('div');
      const viewBoxAttr =
        element.dataset.viewbox ??
        `${metrics.x} ${metrics.y} ${rect.width} ${rect.height}`;
      temp.innerHTML = `<svg width="${rect.width}" height="${rect.height}" viewBox="${viewBoxAttr}" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" overflow="hidden">${element.innerHTML}</svg>`;
      const datauri = HTML2PPTX.svgToDataURI(temp.firstElementChild);
      slide.addImage(
        this.#applyPercentPosition({
          data: datauri,
          x: metrics.x,
          y: metrics.y,
          h: metrics.h,
          w: metrics.w,
          altText: 'Logo ArcelorMittal',
          sizing: { type: 'contain' },
        })
      );
      return;
    }

    if (tag === 'TEXT') {
      const textShape = {
        shape: pptx.ShapeType.rect,
        ...this.#applyPercentPosition({
          x: textMetrics.x - this.#pxToEmu(2.5),
          y: textMetrics.y,
        }),
        h: textMetrics.h,
        w: textMetrics.w + this.#pxToEmu(5),
        breakLine: false,
        color: colors.fill.hex,
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        bold: font.bold,
        margin: 0,
        ...alignment,
      };
      slide.addText(String(element.textContent).trim(), {
        ...textShape,
        ...this.#applyPercentPosition({
          x: this.#alignX(textShape, textMetrics),
        }),
      });
      return;
    }

    if (tag === 'LINE') {
      const linePoints = this.#resolveLinePoints(element, context);
      if (!linePoints) return;
      const lineShape = {
        ...this.#applyPercentPosition(linePoints.start),
        w: linePoints.end.x - linePoints.start.x,
        h: linePoints.end.y - linePoints.start.y,
        ...this.#lineOptions(colors.stroke, borderWidth, dashType),
      };
      slide.addShape(pptx.ShapeType.line, lineShape);
      return;
    }

    if (tag === 'CIRCLE') {
      slide.addShape(
        pptx.ShapeType.ellipse,
        this.#applyPercentPosition({
          x: metrics.x,
          y: metrics.y,
          h: metrics.h,
          w: metrics.w,
          fill: this.#solidFill(colors.fill),
          ...this.#lineOptions(colors.stroke, borderWidth, dashType),
        })
      );
      return;
    }

    if (tag === 'POLYGON' || tag === 'PATH') {
      const temp = document.createElement('div');
      const clone = element.cloneNode(true);
      const svgWidth = Math.max(1, rect.width);
      const svgHeight = Math.max(1, rect.height);
      const adjustedViewBox = this.#applyInverseTransform(
        rect,
        element
      );
      const viewBoxAttr = `${adjustedViewBox.x} ${adjustedViewBox.y} ${Math.max(
        1,
        adjustedViewBox.width
      )} ${Math.max(1, adjustedViewBox.height)}`;
      clone.style.transform = '';
      temp.innerHTML = `<svg width="${svgWidth}" height="${svgHeight}" viewBox="${viewBoxAttr}" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">${clone.outerHTML}</svg>`;
      const datauri = HTML2PPTX.svgToDataURI(temp.firstElementChild);
      slide.addImage(
        this.#applyPercentPosition({
          data: datauri,
          x: metrics.x,
          y: metrics.y,
          h: metrics.h,
          w: metrics.w,
          sizing: { type: 'contain' },
        })
      );
      return;
    }

    if (tag === 'RECT') {
      slide.addShape(
        radius === 0 ? pptx.ShapeType.rect : pptx.ShapeType.roundRect,
        this.#applyPercentPosition({
          x: metrics.x,
          y: metrics.y,
          h: metrics.h,
          w: metrics.w,
          fill: this.#solidFill(colors.fill),
          ...this.#lineOptions(colors.stroke, borderWidth, dashType),
          ...(radius === 0 ? {} : { rectRadius: radius }),
        })
      );
      return;
    }

    if (['DIV', 'TR', 'TD', 'LI'].includes(tag)) {
      const shapeOptions = {
        x: metrics.x,
        y: metrics.y,
        h: metrics.h,
        w: metrics.w,
        fill: this.#solidFill(colors.fill),
        ...this.#lineOptions(colors.stroke, borderWidth, dashType),
        ...(radius === 0 ? {} : { rectRadius: radius }),
      };

      const className = String(element.className);
      if (tag === 'DIV' || tag === 'TD') {
      slide.addShape(
        radius === 0 ? pptx.ShapeType.rect : pptx.ShapeType.roundRect,
        this.#applyPercentPosition(shapeOptions)
      );
      }

      if (className.includes('shape-only')) return;

      const textNodes = element.querySelectorAll(
        '.export-as-text,b,i,u,s'
      );

      if (textNodes.length === 0) {
        const textOptions = this.#paragraphOptions(
          textMetrics,
          font,
          colors.color,
          alignment
        );
        slide.addText(
          String(element.textContent).trim(),
          this.#applyPercentPosition({
            ...textOptions,
            x: this.#alignX(textOptions, textMetrics),
          })
        );
        return;
      }

      for (const node of textNodes) {
        for (const span of node.querySelectorAll('span')) {
          span.removeAttribute('style');
        }
        node.innerHTML = HTML2PPTX.normalizeEntry(node.innerHTML);
        const textObjects = Array.from(node.childNodes).map((child) =>
          this.#textRunFromNode(child, style)
        );
        const bulletPrefix =
          tag === 'LI' ? [{ text: '•   ' }] : [];
        const textOptions = this.#paragraphOptions(
          textMetrics,
          font,
          colors.color,
          alignment
        );
        slide.addText(
          [...bulletPrefix, ...textObjects],
          this.#applyPercentPosition({
            ...textOptions,
            x: this.#alignX(textOptions, textMetrics),
          })
        );
      }
    }
  }

  /**
   * @summary Builds a pptxgen solid fill definition honoring alpha transparency.
   * @private
   * @inner
   * @param {{hex:string,alpha:number}|null} color Fill color descriptor.
   * @returns {{color:string,transparency:number,type:'solid'}} Solid fill options.
   */
 #solidFill(color) {
    return {
      color: color?.hex ?? '#000000',
      transparency: this.#alphaToTransparency(color?.alpha ?? 0),
      type: 'solid',
    };
  }

  /**
   * @summary Builds pptxgen line configuration, supporting transparent strokes.
   * @private
   * @inner
   * @param {{hex:string,alpha:number}|null} color Stroke color definition.
   * @param {number} width Stroke width in EMU.
   * @param {'solid'|'dashDot'|string} dashType Dash style used for pptxgen.
   * @returns {{line:{color:string,width:number,dashType:string,transparency:number}}|{}} Line object or empty object when suppressed.
   */
  #lineOptions(color, width, dashType) {
    if (!Number.isFinite(width) || width <= 0) return {};
    return {
      line: {
        color: color?.hex ?? '#000000',
        width,
        dashType,
        transparency: this.#alphaToTransparency(color?.alpha ?? 1),
      },
    };
  }

  /**
   * @summary Converts normalized alpha (0-1) to pptx transparency percent (0-100).
   * @private
   * @inner
   * @param {number} alpha Alpha channel value (0-1).
   * @returns {number} Transparency percentage accepted by pptxgen.
   */
  #alphaToTransparency(alpha) {
    const safeAlpha = Number.isFinite(alpha) ? alpha : 1;
    return Math.round((1 - Math.max(0, Math.min(1, safeAlpha))) * 100);
  }

  /**
   * @summary Builds default text-box settings mirroring HTML typography.
   * @private
   * @inner
   * @param {{x:number,y:number,w:number,h:number}} metrics Bounding box in EMU.
   * @param {{fontFace:string,fontSize:number,bold:boolean}} font Derived font information.
   * @param {string} color Foreground color.
   * @param {{align:string,autoFit:boolean}} alignment Alignment descriptor.
   * @returns {Object} pptxgen text box options.
   */
  #paragraphOptions(metrics, font, color, alignment) {
    return {
      x: metrics.x,
      y: metrics.y,
      h: metrics.h,
      w: metrics.w,
      color: color?.hex ?? '#000000',
      fontFace: font.fontFace,
      fontSize: font.fontSize,
      bold: font.bold,
      fill: { color: 'FFFFFF', transparency: 99 },
      ...alignment,
      paraSpaceAfter: 0,
      paraSpaceBefore: 0,
      margin: 0,
      fit: 'shrink',
    };
  }

  /**
   * @summary Adjusts metrics by CSS margin and padding to align text content with HTML layout.
   * @private
   * @inner
   * @param {{x:number,y:number,w:number,h:number}} metrics Base metrics in EMU.
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {{x:number,y:number,w:number,h:number}} Metrics adjusted for margin/padding.
   */
  #textBoxMetrics(metrics, style) {
    if (!metrics || !style) return metrics;
    const padding = this.#boxModel(style, 'padding');
    const margin = this.#boxModel(style, 'margin');
    const insetLeft = padding.left + margin.left;
    const insetRight = padding.right + margin.right;
    const insetTop = padding.top + margin.top;
    const insetBottom = padding.bottom + margin.bottom;
    const width = metrics.w - insetLeft - insetRight;
    const height = metrics.h - insetTop - insetBottom;
    if (width <= 0 || height <= 0) {
      return metrics;
    }
    return {
      x: metrics.x + insetLeft,
      y: metrics.y + insetTop,
      w: width,
      h: height,
    };
  }

  /**
   * @summary Applies percent-based coordinates when available, keeping EMU fallbacks otherwise.
   * @private
   * @inner
   * @param {{x?:number|string,y?:number|string,xPercent?:string,yPercent?:string}} options pptxgen position options.
   * @returns {Object} Updated options ready for pptxgen.
   */
  #applyPercentPosition(options) {
    if (!options) return options;
    const next = { ...options };
    if (typeof next.xPercent === 'string') {
      next.x = next.xPercent;
    }
    if (typeof next.yPercent === 'string') {
      next.y = next.yPercent;
    }
    delete next.xPercent;
    delete next.yPercent;
    delete next.viewportPercentX;
    delete next.viewportPercentY;
    return next;
  }

  /**
   * @summary Reads CSS box-model values (margin/padding) converting them to EMU.
   * @private
   * @inner
   * @param {CSSStyleDeclaration} style Computed style.
   * @param {'margin'|'padding'} prop Box-model prefix.
   * @returns {{top:number,right:number,bottom:number,left:number}} Box dimensions in EMU.
   */
  #boxModel(style, prop) {
    return {
      top: this.#pxToEmu(this.#parseLength(style.getPropertyValue(`${prop}-top`))),
      right: this.#pxToEmu(this.#parseLength(style.getPropertyValue(`${prop}-right`))),
      bottom: this.#pxToEmu(this.#parseLength(style.getPropertyValue(`${prop}-bottom`))),
      left: this.#pxToEmu(this.#parseLength(style.getPropertyValue(`${prop}-left`))),
    };
  }

  /**
   * @summary Parses CSS length values returning pixels.
   * @private
   * @inner
   * @param {string} value CSS length.
   * @returns {number} Parsed value in pixels.
   */
  #parseLength(value) {
    const parsed = parseFloat(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  /**
   * @summary Resolves absolute slide coordinates for SVG line endpoints.
   * @private
   * @inner
   * @param {SVGLineElement} element Target line element.
   * @param {{svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}, slideSizing: 'fit'|'viewport'}} context Rendering context.
   * @returns {{start:{x:number,y:number,xPercent?:string,yPercent?:string},end:{x:number,y:number,xPercent?:string,yPercent?:string}}|null} Start/end slide coordinates.
   */
  #resolveLinePoints(element, context) {
    const x1 = this.#parseLength(element.getAttribute('x1'));
    const y1 = this.#parseLength(element.getAttribute('y1'));
    const x2 = this.#parseLength(element.getAttribute('x2'));
    const y2 = this.#parseLength(element.getAttribute('y2'));
    if ([x1, y1, x2, y2].some((value) => !Number.isFinite(value))) return null;
    const start = this.#convertSvgPointToSlide(element, x1, y1, context);
    const end = this.#convertSvgPointToSlide(element, x2, y2, context);
    if (!start || !end) return null;
    return { start, end };
  }

  /**
   * @summary Converts an SVG point (in element coordinates) to slide EMU coordinates.
   * @private
   * @inner
   * @param {SVGElement} element Source element.
   * @param {number} pointX SVG X coordinate.
   * @param {number} pointY SVG Y coordinate.
   * @param {{svgRect: DOMRect, slideSizing: 'fit'|'viewport'}} context Rendering context.
   * @returns {{x:number,y:number,xPercent?:string,yPercent?:string}|null} Slide coordinate descriptor.
   */
  #convertSvgPointToSlide(element, pointX, pointY, context) {
    const svgRoot = element.ownerSVGElement;
    if (
      !svgRoot ||
      typeof svgRoot.createSVGPoint !== 'function' ||
      typeof element.getScreenCTM !== 'function'
    ) {
      return null;
    }
    try {
      const matrix = element.getScreenCTM();
      if (!matrix) return null;
      const svgPoint = svgRoot.createSVGPoint();
      svgPoint.x = pointX;
      svgPoint.y = pointY;
      const domPoint = svgPoint.matrixTransform(matrix);
      const { svgRect } = context;
      const svgX = svgRect.x ?? svgRect.left ?? 0;
      const svgY = svgRect.y ?? svgRect.top ?? 0;
      const safeDomWidth = svgRect.width || 1;
      const safeDomHeight = svgRect.height || 1;
      const ratioX = (domPoint.x - svgX) / safeDomWidth;
      const ratioY = (domPoint.y - svgY) / safeDomHeight;
      const coordinate = {
        x: this.#round(ratioX * this.slideWidthEmu),
        y: this.#round(ratioY * this.slideHeightEmu),
      };
      if (context.slideSizing === 'viewport') {
        if (ratioX < 0 || ratioX > 1) {
          coordinate.xPercent = (ratioX * 100).toFixed(2) + '%';
        }
        if (ratioY < 0 || ratioY > 1) {
          coordinate.yPercent = (ratioY * 100).toFixed(2) + '%';
        }
      }
      return coordinate;
    } catch (error) {
      return null;
    }
  }

  /**
   * @summary Parses CSS transform values to adjust viewBox bounds.
   * @private
   * @inner
   * @param {{x:number,y:number,width:number,height:number}} rect Base rectangle.
   * @param {SVGElement} element Target SVG element.
   * @param {{svgRect: DOMRect}} context Rendering context.
   * @returns {{x:number,y:number,width:number,height:number}|null} Adjusted rect or null when unused.
   */
  #parseTransformRect(rect, element, context) {
    const transform = element.getAttribute('transform') || element.style.transform;
    if (!transform) return null;
    const svgRect = context.svgRect;
    const svgWidth = svgRect.width || 1;
    const svgHeight = svgRect.height || 1;
    let offsetX = rect.x;
    let offsetY = rect.y;
    let width = rect.width;
    let height = rect.height;
    const transforms = transform
      .split(/\)\s*/)
      .map((part) => part.trim())
      .filter(Boolean);
    for (const transformPart of transforms) {
      if (transformPart.startsWith('translate')) {
        const values = transformPart
          .replace(/[a-z]+\(/i, '')
          .replace(')', '')
          .split(/[ ,]+/)
          .map((value) => parseFloat(value));
        offsetX += Number.isFinite(values[0])
          ? values[0]
          : 0;
        offsetY += Number.isFinite(values[1])
          ? values[1]
          : 0;
      } else if (transformPart.startsWith('scale')) {
        const values = transformPart
          .replace(/[a-z]+\(/i, '')
          .replace(')', '')
          .split(/[ ,]+/)
          .map((value) => parseFloat(value));
        const scaleX = Number.isFinite(values[0]) ? values[0] : 1;
        const scaleY = Number.isFinite(values[1]) ? values[1] : scaleX;
        width *= scaleX;
        height *= scaleY;
      }
    }
    return {
      x: (offsetX / svgWidth) * svgRect.width,
      y: (offsetY / svgHeight) * svgRect.height,
      width,
      height,
    };
  }

  /**
   * @summary Applies inverse transform factors to viewBox coordinates.
   * @private
   * @inner
   * @param {{x:number,y:number,width:number,height:number}} viewBoxRect ViewBox rectangle.
   * @param {SVGElement} element Target SVG element.
   * @returns {{x:number,y:number,width:number,height:number}} Corrected viewBox rectangle.
   */
  #applyInverseTransform(viewBoxRect, element) {
    const transform = element.getAttribute('transform') || element.style.transform;
    if (!transform) return viewBoxRect;
    let { x, y, width, height } = viewBoxRect;
    const transforms = transform
      .split(/\)\s*/)
      .map((part) => part.trim())
      .filter(Boolean);
    for (const transformPart of transforms) {
      if (transformPart.startsWith('translate')) {
        const values = transformPart
          .replace(/[a-z]+\(/i, '')
          .replace(')', '')
          .split(/[ ,]+/)
          .map((value) => parseFloat(value));
        x = 0; y = 0;
        x = -1 * (Number.isFinite(values[0]) ? values[0] : 0);
        y = -1 * (Number.isFinite(values[1]) ? values[1] : 0);
      } else if (transformPart.startsWith('scale')) {
        const values = transformPart
          .replace(/[a-z]+\(/i, '')
          .replace(')', '')
          .split(/[ ,]+/)
          .map((value) => parseFloat(value));
        const scaleX = Number.isFinite(values[0]) ? values[0] : 1;
        const scaleY = Number.isFinite(values[1]) ? values[1] : scaleX;
        if (scaleX !== 0) {
          width /= scaleX;
        }
        if (scaleY !== 0) {
          height /= scaleY;
        }
      }
    }
    return { x, y, width, height };
  }


  /**
   * @summary Translates inline DOM nodes into pptxgen text runs.
   * @private
   * @inner
   * @param {Node} node Node to convert.
   * @param {CSSStyleDeclaration} fallbackStyle Parent style to fall back on for text nodes.
   * @returns {{text:string,options:Object}} Text run descriptor.
   */
  #textRunFromNode(node, fallbackStyle) {
    const nodeStyle =
      node.nodeType === Node.TEXT_NODE
        ? fallbackStyle
        : window.getComputedStyle(node);
    const fontWeight = nodeStyle.getPropertyValue('font-weight');
    const fontStyle = nodeStyle.getPropertyValue('font-style');
    const textDecoration = nodeStyle.getPropertyValue('text-decoration');
    return {
      text: node.textContent,
      options: {
        breakLine:
          node.nodeType !== Node.TEXT_NODE &&
          ['DIV', 'P', 'BR'].includes(node.nodeName.toUpperCase()),
        bold: ['bold', 'bolder', '600', '700', '800'].includes(
          fontWeight
        ),
        italic: fontStyle === 'italic',
        strike: textDecoration.includes('line-through'),
        underline: textDecoration.includes('underline'),
      },
    };
  }

  /**
   * @summary Adjusts X coordinate for center/right aligned text to mimic SVG rendering.
   * @private
   * @inner
   * @param {{align?:string,w?:number,x?:number}} options pptxgen text options.
   * @param {{x:number,w:number}} metrics Bounding metrics.
   * @returns {number} Corrected X coordinate in EMU.
   */
  #alignX(options, metrics) {
    if (options.align === 'center') {
      return metrics.x + metrics.w / 2 - options.w / 2;
    }
    if (options.align === 'right') {
      return metrics.x + metrics.w - options.w;
    }
    return options.x ?? metrics.x;
  }

  /**
   * @summary Resolves final fill/stroke/text colors with HTML backgrounds and opacity in mind.
   * @private
   * @inner
   * @param {Element} element Current DOM element.
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {{background:{hex:string,alpha:number},color:{hex:string,alpha:number},fill:{hex:string,alpha:number},stroke:{hex:string,alpha:number}}} Color palette for pptxgen.
   */
  #resolveColors(element, style) {
    const baseOpacity = this.#parseOpacity(style.getPropertyValue('opacity'));
    const background = this.#applyOpacity(
      HTML2PPTX.toColor(style.getPropertyValue('background-color')),
      baseOpacity
    );
    const textColor = this.#applyOpacity(
      HTML2PPTX.toColor(style.getPropertyValue('color')),
      baseOpacity
    );
    const fillColor = this.#applyOpacity(
      HTML2PPTX.toColor(style.getPropertyValue('fill')),
      this.#parseOpacity(style.getPropertyValue('fill-opacity')) * baseOpacity
    );
    const strokeColor = this.#applyOpacity(
      HTML2PPTX.toColor(style.getPropertyValue('stroke')),
      this.#parseOpacity(style.getPropertyValue('stroke-opacity')) * baseOpacity
    );
    const borderColor = this.#applyOpacity(
      HTML2PPTX.toColor(style.getPropertyValue('border-color')),
      baseOpacity
    );
    const isSvgElement = HTML2PPTX.SVG_ELEMENTS.has(element.tagName.toUpperCase());
    const fill = !isSvgElement && this.#hasVisibleColor(background)
      ? background
      : fillColor;
    const stroke = !isSvgElement && this.#hasVisibleColor(borderColor)
      ? borderColor
      : strokeColor;
    return { background, color: textColor, fill, stroke };
  }

  /**
   * @summary Parses CSS opacity values, falling back to `1`.
   * @private
   * @inner
   * @param {string} value CSS opacity value.
   * @returns {number} Normalized opacity between 0 and 1.
   */
  #parseOpacity(value) {
    const parsed = parseFloat(value);
    if (Number.isNaN(parsed)) return 1;
    return Math.max(0, Math.min(1, parsed));
  }

  /**
   * @summary Applies an additional opacity multiplier to a color descriptor.
   * @private
   * @inner
   * @param {{hex:string,alpha:number}} color Base color descriptor.
   * @param {number} multiplier Opacity multiplier (0-1).
   * @returns {{hex:string,alpha:number}} Adjusted color descriptor.
   */
  #applyOpacity(color, multiplier = 1) {
    const safeColor = color ?? { hex: '#000000', alpha: 0 };
    const safeMultiplier = Number.isFinite(multiplier) ? multiplier : 1;
    return {
      hex: safeColor.hex,
      alpha: Math.max(0, Math.min(1, safeColor.alpha * safeMultiplier)),
    };
  }

  /**
   * @summary Indicates whether a color produces a visible result.
   * @private
   * @inner
   * @param {{alpha:number}|null} color Color descriptor.
   * @returns {boolean} True when the color is visible.
   */
  #hasVisibleColor(color) {
    return Boolean(color && color.alpha > 0);
  }

  /**
   * @summary Resolves stroke width in EMU from CSS values.
   * @private
   * @inner
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {number} Stroke width expressed in EMU.
   */
  #resolveBorderWidth(style) {
    const strokeWidth = parseFloat(
      style.getPropertyValue('stroke-width')
    );
    const borderWidth = parseFloat(
      style.getPropertyValue('border-width')
    );
    const pxWidth = [strokeWidth, borderWidth].find(
      (value) => Number.isFinite(value) && value > 0
    );
    if (!Number.isFinite(pxWidth) || pxWidth <= 0) return 0;
    const emuWidth = this.#pxToEmu(pxWidth);
    if (emuWidth <= HTML2PPTX.STROKE_LIMIT) return 0;
    const pptxWidth = Math.round(this.#pxToPoints(pxWidth));
    return Math.max(1, Math.min(256, pptxWidth));
  }

  /**
   * @summary Extracts font family, size (pt) and weight.
   * @private
   * @inner
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {{fontFace:string,fontSize:number,bold:boolean}} Font descriptor.
   */
  #resolveFont(style) {
    const family = String(
      style.getPropertyValue('font-family')
    )
      .split(',')[0]
      .replace(/["']/g, '')
      .trim();
    const fontSizePx = parseFloat(style.getPropertyValue('font-size'));
    const fontWeight = style.getPropertyValue('font-weight');
    return {
      fontFace: family,
      fontSize: this.#pxToPoints(fontSizePx),
      bold: ['bold', 'bolder', '600', '700', '800'].includes(fontWeight),
    };
  }

  /**
   * @summary Converts CSS/SVG radius definitions to pptx round-rect ratios.
   * @private
   * @inner
   * @param {Element} element DOM element being inspected.
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @param {{width:number,height:number}} rect Bounding rectangle in pixels.
   * @returns {number} Radius ratio between 0 and 1.
   */
  #resolveRadius(element, style, rect) {
    const rx = parseFloat(element.getAttribute('rx'));
    const ry = parseFloat(element.getAttribute('ry'));
    const borderRadius = parseFloat(
      style.getPropertyValue('border-radius')
    );
    const pxRadius =
      !Number.isNaN(borderRadius) && borderRadius > 0
        ? borderRadius
        : Number.isNaN(rx) || Number.isNaN(ry)
        ? 0
        : (rx + ry) / 2;
    if (!pxRadius || pxRadius <= 0) return 0;
    const width = Number.isFinite(rect?.width) ? rect.width : pxRadius * 2;
    const height = Number.isFinite(rect?.height) ? rect.height : pxRadius * 2;
    const reference = 1//Math.max(1, Math.min(width, height));
    const normalized = pxRadius / (reference / 2);
    return Math.max(0, Math.min(1, normalized)) * 0.05;
  }

  /**
   * @summary Formats the current date using the configured locale.
   * @private
   * @inner
   * @returns {string} Localized date string.
   */
  #formatDate() {
    try {
      return new Date().toLocaleDateString(this.locale);
    } catch (error) {
      return new Date().toLocaleDateString();
    }
  }

  /**
   * @summary Applies the same visibility rules as the editor to skip certain nodes.
   * @private
   * @inner
   * @param {Element} element DOM element to inspect.
   * @returns {boolean} True when the node should be skipped.
   */
  #shouldSkipElement(element) {
    const className = String(element.className);
    if (className.includes('no-export')) return true;
    if (className.includes('hide-on-export')) return true;
    if (
      className.includes('hide-on-presentation') &&
      this.editor &&
      this.editor.classList.contains('presentation')
    ) {
      return true;
    }
    return false;
  }

  /**
   * @summary Converts pixels to point units used by pptxgen.
   * @private
   * @inner
   * @param {number} px Pixel value.
   * @returns {number} Equivalent in points.
   */
  #pxToPoints(px) {
    const value = Number.isFinite(px) ? px : 0;
    return (value * 72) / HTML2PPTX.PX_PER_IN;
  }

  /**
   * @summary Converts pixels to EMU using the class conversion factor.
   * @private
   * @inner
   * @param {number} px Pixel value.
   * @returns {number} Equivalent in EMU.
   */
  #pxToEmu(px) {
    const value = Number.isFinite(px) ? px : 0;
    return this.#round(value * HTML2PPTX.PX_TO_EMU);
  }

  /**
   * @summary Applies safe rounding to avoid floating point artifacts.
   * @private
   * @inner
   * @param {number} value Numeric value.
   * @returns {number} Rounded integer.
   */
  #round(value) {
    return Math.round((value ?? 0) + Number.EPSILON);
  }

  /**
   * @summary Reads and normalizes SVG viewBox data.
   * @private
   * @inner
   * @param {SVGElement} svg Source SVG element.
   * @param {DOMRect} rect Bounding client rect for fallback dimensions.
   * @returns {{minX:number,minY:number,width:number,height:number}} Normalized viewBox object.
   */
  #getViewBox(svg, rect) {
    const raw = svg.getAttribute('viewBox');
    if (!raw) {
      return { minX: 0, minY: 0, width: rect.width, height: rect.height };
    }
    const parts = raw
      .trim()
      .split(/[ ,]+/)
      .map((part) => Number(part));
    if (parts.length === 4 && parts.every((part) => Number.isFinite(part))) {
      return { minX: parts[0], minY: parts[1], width: parts[2], height: parts[3] };
    }
    return { minX: 0, minY: 0, width: rect.width, height: rect.height };
  }

  /**
   * @summary Converts a DOMRect to slide EMU coordinates honoring ratios and render size.
   * @private
   * @inner
   * @param {{x:number,y:number,width:number,height:number,coordinateSpace?:'svg'|'screen'}} rect Bounding rectangle.
   * @param {{svgRect: DOMRect, viewBox: {width:number,height:number}, slideSizing: 'fit'|'viewport'}} context Precomputed slide context.
   * @returns {{x:number,y:number,w:number,h:number}} Dimensions translated to EMU.
   */
  #rectToSlideMetrics(rect, context) {
    const { svgRect, viewBox, slideSizing } = context;
    const safeDomWidth = svgRect.width || 1;
    const safeDomHeight = svgRect.height || 1;
    const rectX = rect.x ?? rect.left ?? 0;
    const rectY = rect.y ?? rect.top ?? 0;
    const svgX = svgRect.x ?? svgRect.left ?? 0;
    const svgY = svgRect.y ?? svgRect.top ?? 0;
    const rectWidth =
      rect.width ??
      Math.max(0, (rect.right ?? rectX) - (rect.left ?? rectX));
    const rectHeight =
      rect.height ??
      Math.max(0, (rect.bottom ?? rectY) - (rect.top ?? rectY));
    const rectSpace = rect.coordinateSpace ?? 'screen';

    const safeViewWidth = viewBox.width || safeDomWidth || 1;
    const safeViewHeight = viewBox.height || safeDomHeight || 1;
    const viewMinX = viewBox.minX ?? 0;
    const viewMinY = viewBox.minY ?? 0;

    if (slideSizing === 'viewport') {
      const viewWidth = viewBox.width || safeDomWidth || 1;
      const viewHeight = viewBox.height || safeDomHeight || 1;
      let viewX;
      let viewY;
      let viewW;
      let viewH;
      if (rectSpace === 'svg') {
        viewX = rectX;
        viewY = rectY;
        viewW = rectWidth;
        viewH = rectHeight;
      } else {
        const domToViewBoxX = (value) =>
          (value / safeDomWidth) * viewWidth;
        const domToViewBoxY = (value) =>
          (value / safeDomHeight) * viewHeight;
        viewX = domToViewBoxX(rectX - svgX) + viewMinX;
        viewY = domToViewBoxY(rectY - svgY) + viewMinY;
        viewW = domToViewBoxX(rectWidth);
        viewH = domToViewBoxY(rectHeight);
      }
      const percentX = (viewX - viewMinX) / viewWidth;
      const percentY = (viewY - viewMinY) / viewHeight;
      return {
        x: this.#round(percentX * this.slideWidthEmu),
        y: this.#round(percentY * this.slideHeightEmu),
        w: this.#round((viewW / viewWidth) * this.slideWidthEmu),
        h: this.#round((viewH / viewHeight) * this.slideHeightEmu),
        viewportPercentX: percentX,
        viewportPercentY: percentY,
      };
    }

    let viewX;
    let viewY;
    let viewW;
    let viewH;
    if (rectSpace === 'svg') {
      viewX = rectX - viewMinX;
      viewY = rectY - viewMinY;
      viewW = rectWidth;
      viewH = rectHeight;
    } else {
      const domToViewBoxX = (value) =>
        (value / safeDomWidth) * safeViewWidth;
      const domToViewBoxY = (value) =>
        (value / safeDomHeight) * safeViewHeight;
      viewX = domToViewBoxX(rectX - svgX);
      viewY = domToViewBoxY(rectY - svgY);
      viewW = domToViewBoxX(rectWidth);
      viewH = domToViewBoxY(rectHeight);
    }

    return {
      x: this.#round((viewX / safeViewWidth) * this.slideWidthEmu),
      y: this.#round((viewY / safeViewHeight) * this.slideHeightEmu),
      w: this.#round((viewW / safeViewWidth) * this.slideWidthEmu),
      h: this.#round((viewH / safeViewHeight) * this.slideHeightEmu),
    };
  }

  /**
   * @summary Converts a DOMRect into the original SVG viewBox coordinate system.
   * @private
   * @inner
   * @param {{x:number,y:number,width:number,height:number,coordinateSpace?:'svg'|'screen'}} rect Bounding rectangle.
   * @param {{svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}}} context Slide context.
   * @returns {{x:number,y:number,width:number,height:number}} Rectangle expressed in viewBox units.
   */
  #rectToViewBoxRect(rect, context) {
    if (rect.coordinateSpace === 'svg') {
      return {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
      };
    }
    const { svgRect, viewBox } = context;
    const safeDomWidth = svgRect.width || 1;
    const safeDomHeight = svgRect.height || 1;
    const rectX = rect.x ?? rect.left ?? 0;
    const rectY = rect.y ?? rect.top ?? 0;
    const svgX = svgRect.x ?? svgRect.left ?? 0;
    const svgY = svgRect.y ?? svgRect.top ?? 0;
    const rectWidth =
      rect.width ??
      Math.max(0, (rect.right ?? rectX) - (rect.left ?? rectX));
    const rectHeight =
      rect.height ??
      Math.max(0, (rect.bottom ?? rectY) - (rect.top ?? rectY));
    const viewWidth = viewBox.width || safeDomWidth;
    const viewHeight = viewBox.height || safeDomHeight;
    const domToViewBoxX = (value) =>
      (value / safeDomWidth) * viewWidth;
    const domToViewBoxY = (value) =>
      (value / safeDomHeight) * viewHeight;
    const viewX = domToViewBoxX(rectX - svgX) + (viewBox.minX ?? 0);
    const viewY = domToViewBoxY(rectY - svgY) + (viewBox.minY ?? 0);
    const viewW = domToViewBoxX(rectWidth);
    const viewH = domToViewBoxY(rectHeight);
    return { x: viewX, y: viewY, width: viewW, height: viewH };
  }

  /**
   * @summary Normalizes slide metrics to ensure non-negative coordinates for PPTX export.
   * @private
   * @inner
   * @param {{x:number,y:number,w:number,h:number}} metrics Raw metrics.
   * @param {{slideSizing: 'fit'|'viewport'}} context Rendering context.
   * @returns {{x:number,y:number,w:number,h:number}} Normalized metrics.
   */
  #normalizeMetrics(metrics, context) {
    if (context.slideSizing !== 'viewport') return metrics;
    const normalized = { ...metrics };
    if (metrics.x < 0 && typeof metrics.viewportPercentX === 'number') {
      normalized.xPercent = (metrics.viewportPercentX * 100).toFixed(2) + '%';
      normalized.x = 0;
    }
    if (metrics.y < 0 && typeof metrics.viewportPercentY === 'number') {
      normalized.yPercent = (metrics.viewportPercentY * 100).toFixed(2) + '%';
      normalized.y = 0;
    }
    return normalized;
  }

  /**
   * @summary Attempts to capture transformed bounds using `getBBox` + `getScreenCTM`.
   * @private
   * @inner
   * @param {Element} element Element to inspect.
   * @returns {{x:number,y:number,width:number,height:number,coordinateSpace:'screen'}|null} Transformed rectangle or `null` when unavailable.
   */
  #getClientRect(element) {
    if (
      typeof element.getBBox === 'function' &&
      typeof element.getScreenCTM === 'function' &&
      typeof DOMPoint === 'function'
    ) {
      try {
        const bbox = element.getBBox();
        const matrix = element.getScreenCTM();
        if (!matrix) return null;
        const topLeft = new DOMPoint(bbox.x, bbox.y).matrixTransform(matrix);
        const bottomRight = new DOMPoint(
          bbox.x + bbox.width,
          bbox.y + bbox.height
        ).matrixTransform(matrix);
        return {
          x: topLeft.x,
          y: topLeft.y,
          width: bottomRight.x - topLeft.x,
          height: bottomRight.y - topLeft.y,
          coordinateSpace: 'screen',
        };
      } catch (error) {
        return null;
      }
    }
    return null;
  }

  /**
   * @summary Normalizes inline HTML by stripping problematic tags/styles before export.
   * @static
   * @param {string} value HTML string to sanitize.
   * @returns {string} Sanitized HTML string.
   */
  static normalizeEntry(value) {
    return String(value)
      .replace(/\u0007/gm, '')
      .replace(/<p\b([^>]*?)\s([^>]*)>/gm, '<p>')
      .replace(/<div\b([^>]*?)\s([^>]*)>/gm, '<div>')
      .replace(/<span\b([^>]*?)\s([^>]*)>/gm, '<span>')
      .replace(/<font\b([^>]*?)\s([^>]*)>/gm, '<font>')
      .replace(/ style="[^"]*"/gm, '');
  }

  /**
   * @summary Serializes an SVG element to a base64 data URI for embedding.
   * @static
   * @param {SVGElement} svgElement SVG node to serialize.
   * @returns {string} data URI string.
   */
  static svgToDataURI(svgElement) {
    const svgString = new XMLSerializer().serializeToString(svgElement);
    const encoded = encodeURIComponent(svgString)
      .replace(/'/g, '%27')
      .replace(/"/g, '%22');
    const base64 = btoa(unescape(encoded));
    return `data:image/svg+xml;base64,${base64}`;
  }

  /**
   * @summary Converts any CSS color value to a hex+alpha descriptor.
   * @static
   * @param {string} color CSS color value.
   * @returns {{hex:string,alpha:number}} Color descriptor.
   */
  static toColor(color) {
    try {
      if (!color) return { hex: '#000000', alpha: 0 };
      const ctx = HTML2PPTX.#colorContext();
      ctx.clearRect(0, 0, 1, 1);
      ctx.fillStyle = color;
      ctx.fillRect(0, 0, 1, 1);
      const [r, g, b, a] = ctx.getImageData(0, 0, 1, 1).data;
      const hex =
        '#' +
        [r, g, b]
          .map((channel) => channel.toString(16).padStart(2, '0'))
          .join('');
      return { hex, alpha: Math.max(0, Math.min(1, a / 255)) };
    } catch (error) {
      return { hex: '#000000', alpha: 0 };
    }
  }

  /**
   * @summary Backwards compatible helper returning only the hexadecimal color.
   * @static
   * @param {string} color CSS color value.
   * @returns {string} Hexadecimal color (#RRGGBB).
   */
  static toHex(color) {
    return HTML2PPTX.toColor(color).hex;
  }

  /**
   * @summary Returns the cached 2D canvas context used for color conversions.
   * @private
   * @static
   * @returns {CanvasRenderingContext2D} 1×1 canvas context.
   */
  static #colorContext() {
    if (!this.colorCanvas) {
      this.colorCanvas = document.createElement('canvas');
      this.colorCanvas.width = 1;
      this.colorCanvas.height = 1;
      this.colorCtx = this.colorCanvas.getContext('2d');
    }
    return this.colorCtx;
  }

  /**
   * @summary Replicates the UI visibility check to skip hidden/transparent nodes.
   * @static
   * @param {Element|null} node Element to inspect.
   * @returns {boolean} True when the node (or ancestor) is hidden.
   */
  static isHidden(node) {
    while (node) {
      const style = window.getComputedStyle(node);
      if (
        style.display === 'none' ||
        style.visibility === 'hidden' ||
        style.opacity === '0'
      ) {
        return true;
      }
      node = node.parentElement;
    }
    return false;
  }

  /**
   * @summary Pixels-per-inch reference used by browsers.
   * @constant
   * @returns {number} Pixels per inch.
   */
  static get PX_PER_IN() {
    return 96;
  }

  /**
   * @summary Number of EMUs per inch defined by the PPTX format.
   * @constant
   * @returns {number} EMUs per inch.
   */
  static get EMU_PER_IN() {
    return 914400;
  }

  /**
   * @summary Conversion factor from pixels to EMUs.
   * @constant
   * @returns {number} EMUs per pixel.
   */
  static get PX_TO_EMU() {
    return HTML2PPTX.EMU_PER_IN / HTML2PPTX.PX_PER_IN;
  }
}

const _html2pptxGlobal =
  typeof globalThis !== 'undefined'
    ? globalThis
    : typeof window !== 'undefined'
    ? window
    : typeof global !== 'undefined'
    ? global
    : this;

if (typeof define === 'function' && define.amd) {
  define(() => HTML2PPTX);
}

if (typeof module === 'object' && module.exports) {
  module.exports = HTML2PPTX;
  module.exports.default = HTML2PPTX;
  module.exports.HTML2PPTX = HTML2PPTX;
  Object.defineProperty(module.exports, '__esModule', { value: true });
} else if (_html2pptxGlobal) {
  _html2pptxGlobal.HTML2PPTX = HTML2PPTX;
}
