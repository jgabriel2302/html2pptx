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
 * @description Works as either an ES module (`import HTML2PPTX from './SVGToPPTX.js';`) or a classic `<script>` that exposes a global constructor.
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
    return 0.013683353027016402;
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
  }

  /**
   * @summary Converts DOM slides into PPTX slides.
   * @description Iterates over SVG nodes, converting each to pptxgen shapes/text. Can append to an existing presentation.
   * @param {Iterable<SVGElement>|SVGElement|null} slidesSvg Collection of SVG elements or a single SVG node.
   * @param {PptxGenJS|null} [recycle=null] Existing pptxgen instance used to append slides; when omitted a new instance is created and written to disk.
   * @returns {PptxGenJS} The pptxgen presentation instance, useful when chaining additional operations.
   */
  generate(slidesSvg, recycle = null) {
    const pptx = recycle ?? this.#createPresentation();
    const nodes = this.#normalizeSlides(slidesSvg);
    for (const svg of nodes) {
      this.#renderSlide(pptx, svg);
    }
    if (!recycle) {
      pptx.writeFile({ fileName: this.fileName });
    }
    return pptx;
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
   * @summary Creates a PPTX slide and processes relevant SVG children preserving the aspect ratio.
   * @private
   * @inner
   * @param {PptxGenJS} pptx Target pptxgen presentation.
   * @param {SVGElement} svg Source SVG element representing a slide.
   * @returns {void}
   */
  #renderSlide(pptx, svg) {
    const slide = pptx.addSlide();
    const svgRect = svg.getBoundingClientRect();
    const viewBox = this.#getViewBox(svg, svgRect);
    const context = { svgRect, viewBox, slide };
    const elements = svg.querySelectorAll(
      'text,rect,g[name],div,li,td'
    );
    for (const element of elements) {
      this.#renderElement(element, context);
    }
  }

  /**
   * @summary Converts a single DOM element into pptxgen shapes, text or images.
   * @private
   * @inner
   * @param {Element} element DOM element being exported.
   * @param {{svgRect: DOMRect, viewBox: {minX:number,minY:number,width:number,height:number}, slide: PptxGenJS.Slide}} context Precomputed slide context.
   * @returns {void}
   */
  #renderElement(element, context) {
    if (this.#shouldSkipElement(element)) return;

    const rect = this.#getClientRect(element) ?? element.getBoundingClientRect();
    const metrics = this.#rectToSlideMetrics(rect, context);
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
    const radius = this.#resolveRadius(element, style);

    const tag = element.tagName.toUpperCase();
    const name = element.getAttribute('name');
    const slide = context.slide;

    if (name === '@updateDate') {
      const dateOptions = {
        x: metrics.x,
        y: metrics.y,
        h: metrics.h,
        w: metrics.w * 2,
        breakLine: false,
        color: colors.fill,
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        bold: font.bold,
        ...alignment,
      };
      slide.addText(this.#formatDate(), {
        ...dateOptions,
        x: this.#alignX(dateOptions, metrics),
      });
      return;
    }

    if (name === '@pageNumber') {
      slide.slideNumber = {
        x: metrics.x,
        y: metrics.y,
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        color: colors.fill,
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
      slide.addImage({
        data: datauri,
        x: metrics.x,
        y: metrics.y,
        h: metrics.h,
        w: metrics.w,
        altText: 'Logo ArcelorMittal',
        sizing: { type: 'contain' },
      });
      return;
    }

    if (tag === 'TEXT') {
      const textShape = {
        shape: pptx.ShapeType.rect,
        x: metrics.x,
        y: metrics.y,
        h: metrics.h,
        w: metrics.w,
        breakLine: false,
        color: colors.fill,
        fontFace: font.fontFace,
        fontSize: font.fontSize,
        bold: font.bold,
        ...alignment,
      };
      slide.addText(String(element.textContent).trim(), {
        ...textShape,
        x: this.#alignX(textShape, metrics),
      });
      return;
    }

    if (tag === 'RECT') {
      slide.addShape(
        radius === 0 ? pptx.ShapeType.rect : pptx.ShapeType.roundRect,
        {
          x: metrics.x,
          y: metrics.y,
          h: metrics.h,
          w: metrics.w,
          fill: {
            color: colors.fill,
            transparency: colors.fill === '#000000' ? 100 : 0,
            type: 'solid',
          },
          ...this.#lineOptions(colors.stroke, borderWidth, dashType),
          ...(radius === 0 ? {} : { rectRadius: radius }),
        }
      );
      return;
    }

    if (['DIV', 'TR', 'TD', 'LI'].includes(tag)) {
      const shapeOptions = {
        x: metrics.x,
        y: metrics.y,
        h: metrics.h,
        w: metrics.w,
        fill: { color: colors.fill },
        ...this.#lineOptions(colors.stroke, borderWidth, dashType),
        ...(radius === 0 ? {} : { rectRadius: radius }),
      };

      const className = String(element.className);
      if (tag === 'DIV' || tag === 'TD') {
        slide.addShape(
          radius === 0 ? pptx.ShapeType.rect : pptx.ShapeType.roundRect,
          shapeOptions
        );
      }

      if (className.includes('shape-only')) return;

      const textNodes = element.querySelectorAll(
        '.export-as-text,b,i,u,s'
      );

      if (textNodes.length === 0) {
        const textOptions = this.#paragraphOptions(
          metrics,
          font,
          colors.color,
          alignment
        );
        slide.addText(String(element.textContent).trim(), {
          ...textOptions,
          x: this.#alignX(textOptions, metrics),
        });
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
          metrics,
          font,
          colors.color,
          alignment
        );
        slide.addText(
          [...bulletPrefix, ...textObjects],
          {
            ...textOptions,
            x: this.#alignX(textOptions, metrics),
          }
        );
      }
    }
  }

  /**
   * @summary Builds pptxgen line configuration.
   * @description Ignores stroke data when width falls below {@link HTML2PPTX.STROKE_LIMIT}.
   * @private
   * @inner
   * @param {string} color Hex stroke color.
   * @param {number} width Stroke width in EMU.
   * @param {'solid'|'dashDot'|string} dashType Dash style used for pptxgen.
   * @returns {{line:{color:string,width:number,dashType:string}}|{}} Line object or empty object when suppressed.
   */
  #lineOptions(color, width, dashType) {
    if (!width || width <= HTML2PPTX.STROKE_LIMIT) return {};
    return { line: { color, width, dashType } };
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
      color,
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
   * @summary Resolves final fill/stroke/text colors with HTML backgrounds in mind.
   * @private
   * @inner
   * @param {Element} element Current DOM element.
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {{background:string,color:string,fill:string,stroke:string}} Color palette for pptxgen.
   */
  #resolveColors(element, style) {
    const background = HTML2PPTX.toHex(
      style.getPropertyValue('background-color')
    );
    const textColor = HTML2PPTX.toHex(
      style.getPropertyValue('color')
    );
    const fill =
      background !== '#000000' &&
      !HTML2PPTX.SVG_ELEMENTS.has(element.tagName.toUpperCase())
        ? background
        : HTML2PPTX.toHex(style.getPropertyValue('fill'));
    const borderColor = HTML2PPTX.toHex(
      style.getPropertyValue('border-color')
    );
    const stroke =
      borderColor !== '#000000' &&
      !HTML2PPTX.SVG_ELEMENTS.has(element.tagName.toUpperCase())
        ? borderColor
        : HTML2PPTX.toHex(style.getPropertyValue('stroke'));
    return { background, color: textColor, fill, stroke };
  }

  /**
   * @summary Resolves stroke width in EMU from CSS values.
   * @private
   * @inner
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {number} Stroke width expressed in EMU.
   */
  #resolveBorderWidth(style) {
    const borderWidth = parseFloat(
      style.getPropertyValue('border-width')
    );
    const strokeWidth = parseFloat(
      style.getPropertyValue('stroke-width')
    );
    const pxWidth =
      !Number.isNaN(borderWidth) && borderWidth > 0
        ? borderWidth
        : strokeWidth;
    return this.#pxToEmu(pxWidth);
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
   * @summary Converts CSS/SVG radius definitions to EMU to keep rounded corners consistent.
   * @private
   * @inner
   * @param {Element} element DOM element being inspected.
   * @param {CSSStyleDeclaration} style Computed style for the element.
   * @returns {number} Radius in EMU.
   */
  #resolveRadius(element, style) {
    const rx = parseFloat(element.getAttribute('rx'));
    const ry = parseFloat(element.getAttribute('ry'));
    const borderRadius = parseFloat(
      style.getPropertyValue('border-radius')
    );
    const pxRadius = !Number.isNaN(borderRadius) && borderRadius > 0
      ? borderRadius
      : Number.isNaN(rx) || Number.isNaN(ry)
      ? 0
      : (rx + ry) / 2;
    return this.#pxToEmu(pxRadius * 0.75);
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
   * @param {{x:number,y:number,width:number,height:number}} rect Bounding rectangle in client pixels.
   * @param {{svgRect: DOMRect, viewBox: {width:number,height:number}}} context Precomputed slide context.
   * @returns {{x:number,y:number,w:number,h:number}} Dimensions translated to EMU.
   */
  #rectToSlideMetrics(rect, context) {
    const { svgRect, viewBox } = context;
    const safeWidth = svgRect.width || 1;
    const safeHeight = svgRect.height || 1;
    const viewWidth = viewBox.width || safeWidth;
    const viewHeight = viewBox.height || safeHeight;

    const domToViewBoxX = (value) =>
      (value / safeWidth) * viewWidth;
    const domToViewBoxY = (value) =>
      (value / safeHeight) * viewHeight;

    const viewX = domToViewBoxX(rect.x - svgRect.x);
    const viewY = domToViewBoxY(rect.y - svgRect.y);
    const viewW = domToViewBoxX(rect.width);
    const viewH = domToViewBoxY(rect.height);

    return {
      x: this.#round((viewX / viewWidth) * this.slideWidthEmu),
      y: this.#round((viewY / viewHeight) * this.slideHeightEmu),
      w: this.#round((viewW / viewWidth) * this.slideWidthEmu),
      h: this.#round((viewH / viewHeight) * this.slideHeightEmu),
    };
  }

  /**
   * @summary Attempts to capture transformed bounds using `getBBox` + `getScreenCTM`.
   * @private
   * @inner
   * @param {Element} element Element to inspect.
   * @returns {{x:number,y:number,width:number,height:number}|null} Transformed rectangle or `null` when unavailable.
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
   * @summary Converts any CSS color value to an absolute hexadecimal string.
   * @static
   * @param {string} color CSS color value.
   * @returns {string} Hexadecimal color (#RRGGBB).
   */
  static toHex(color) {
    if (!color || color === 'transparent') return '#000000';
    const ctx = HTML2PPTX.#colorContext();
    ctx.clearRect(0, 0, 1, 1);
    ctx.fillStyle = color;
    ctx.fillRect(0, 0, 1, 1);
    const [r, g, b] = ctx.getImageData(0, 0, 1, 1).data;
    return (
      '#' +
      [r, g, b]
        .map((channel) => channel.toString(16).padStart(2, '0'))
        .join('')
    );
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
