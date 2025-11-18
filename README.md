# HTML2PPTX

Helper utilities for exporting any SVG/HTML layout to Microsoft PowerPoint using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/). The class was designed to work as a single drop-in file that runs in browsers (via `<script>`), CommonJS (`require`) or ESM (`import`) contexts without additional bundling.

## Features

- Recreates rendered SVG/HTML tables into native PPTX shapes and text.
- Handles logos/images, page numbers and custom placeholders (e.g., `@updateDate`, `@Logo`, `@Svg`).
- Converts CSS units to EMU/PT with calibrated rounding to keep sizes consistent.
- Skips hidden DOM nodes (`no-export`, `hide-on-export`, `hide-on-presentation`) to mirror the on-screen preview.
- Includes a UMD-style footer so the same file can be used everywhere.

## Getting Started

### Browser (script tag)

```html
<script src="pptxgen.bundle.js"></script>
<script src="SVGToPPTX.js"></script>
<script>
  const exporter = new HTML2PPTX({
    author: 'Me',
    company: 'MyCompany',
    title: 'Report',
    locale: 'pt-BR'
  });

  const slides = document.querySelectorAll('#result svg');
  exporter.generate(slides); // triggers download
</script>
```

Ensure that `PptxGenJS` is loaded **before** `SVGToPPTX.js`.

### CommonJS / Node (require)

```js
const HTML2PPTX = require('./SVGToPPTX.js');
const exporter = new HTML2PPTX({ title: 'Report' });

// Use jsdom/puppeteer to render your SVG before calling:
exporter.generate(svgNodes, existingPptx); // returns the pptxgen instance
```

### ES Modules / Bundlers

```js
import HTML2PPTX from './SVGToPPTX.js';

const exporter = new HTML2PPTX({ title: 'Report' });
exporter.generate(document.querySelectorAll('#result svg'));
```

## API Highlights

- `new HTML2PPTX(options)` – configure author, company, title, subject, locale, editor container, and PPTX layout.
- `generate(slidesSvg, recycle?)` – converts NodeList/arrays or a single SVG element into PPTX slides. Pass an existing `PptxGenJS` instance via `recycle` to append slides without downloading yet.
- Static helpers such as `HTML2PPTX.normalizeEntry()`, `HTML2PPTX.svgToDataURI()`, and `HTML2PPTX.toHex()` are available if you want to reuse the conversion logic elsewhere.

## Reserved CSS Classes

Some classes are interpreted by the exporter to mimic the on-screen editor:

- `no-export` – element is ignored entirely when building the PPTX.
- `hide-on-export` – same as above, but allows the element to remain visible on screen for debugging.
- `hide-on-presentation` – element only appears during editing; when the preview container has the `presentation` class it is skipped.
- `shape-only` – draw the surrounding rectangle but do not emit any text for that element.
- `export-as-text` – forces the exporter to treat the inner node as pure text, bypassing nested spans and inline styles.

Use these classes sparingly so that the generated slides stay in sync with the SVG view.

## Reserved Placeholders

The default template also features placeholders that trigger special behavior:

- `@updateDate` – replaced with the current date formatted using the configured locale.
- `@pageNumber` – turned into the slide number using pptxgen’s native pagination feature.
- `@Logo` / `@Svg` – the inner SVG is rasterized and inserted as an image (use `data-viewbox` to control cropping).

Adding your own placeholders is straightforward: update the template SVG and extend the logic where slides are mounted so those tokens are replaced with actual content before exporting.

## Development Notes

- The repository intentionally keeps everything inside `SVGToPPTX.js` for easy embedding.
- Unit tests are not included; validate exported presentations manually when changing conversion logic (especially px→EMU calculations).
- When modifying the class, keep the JSDoc comments accurate—they power IDE hints for all entry points.

## License

MIT © João Gabriel Corrêa da Silva
