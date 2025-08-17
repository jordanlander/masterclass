/**
 * Build a PPTX presentation from a structured slide model.
 */

export interface SlideElement {
  type: 'title' | 'text' | 'image' | 'footer';
  text?: string;
  src?: string;
  options?: any;
}

// A slide can be a list of structured elements or a pre-rendered image
export interface SlideModel {
  elements?: SlideElement[];
  src?: string; // data URL of a rendered slide image
  notes?: string[];
}

// PptxGenJS is loaded globally via <script>
declare const PptxGenJS: any;

/**
 * Parse simple markdown for **bold** segments and return a value suitable for PptxGenJS addText.
 * If no markdown is present the original string is returned.
 */
function parseBold(text: string): any {
  const regex = /\*\*(.*?)\*\*/g;
  let lastIndex = 0;
  const parts: any[] = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push({ text: text.slice(lastIndex, match.index) });
    }
    parts.push({ text: match[1], options: { bold: true } });
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) {
    parts.push({ text: text.slice(lastIndex) });
  }
  return parts.length ? parts : text;
}

/**
 * Convert an array of slide models into a PptxGenJS presentation.
 * @param slides Array of slide definitions
 * @param meta Optional metadata such as title
 */
export function buildPptx(slides: SlideModel[], meta: { title?: string } = {}): any {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  if (meta.title) {
    pptx.coreProps = { title: meta.title };
  }

  // Attempt to read theme colors from CSS variables; fall back to defaults
  let brand = '#1e3a8a';
  let accent = '#f97316';
  if (typeof window !== 'undefined') {
    const css = getComputedStyle(document.documentElement);
    brand = css.getPropertyValue('--brand').trim() || brand;
    accent = css.getPropertyValue('--accent').trim() || accent;
  }

  const titleBarH = 0.094;
  const accentBarH = 0.031;
  const footerBarH = 0.3125;

  slides.forEach(slideModel => {
    const slide = pptx.addSlide();

    // Top title and accent bars
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: titleBarH, fill: { color: brand }, line: { color: brand } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: titleBarH, w: 10, h: accentBarH, fill: { color: accent }, line: { color: accent } });
    // Bottom accent and footer bars
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.625 - footerBarH - accentBarH, w: 10, h: accentBarH, fill: { color: accent }, line: { color: accent } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.625 - footerBarH, w: 10, h: footerBarH, fill: { color: brand }, line: { color: brand } });

    if (slideModel.src) {
      // slide provided as a full-size image (e.g., html2canvas render)
      slide.addImage({ data: slideModel.src, x: 0, y: 0, w: 10, h: 5.625 });
    } else if (slideModel.elements) {
      let y = 0.5;
      slideModel.elements.forEach(el => {
        switch (el.type) {
          case 'title': {
            const titleText = (el.text || '').replace(/\*\*(.*?)\*\*/g, '$1');
            const options: any = { x: 0.5, y, w: 9, h: 1, fontSize: 32, bold: true, ...(el.options || {}) };
            slide.addText(titleText, options);
            y += options.h;
            break;
          }
          case 'text': {
            const rawText = el.text || '';
            if (el.options && typeof el.options.y === 'number') {
              y = el.options.y;
            }
            const lines = rawText.split('\n').length;
            const options: any = { x: 0.5, y, w: 9, h: 0.6 * lines, fontSize: 18, ...(el.options || {}) };
            const formatted = parseBold(rawText);
            slide.addText(formatted, options);
            y += options.h;
            break;
          }
          case 'image': {
            const options: any = { data: el.src, x: 0.5, y, w: 4, h: 3, ...(el.options || {}) };
            slide.addImage(options);
            y += options.h + 0.5;
            break;
          }
          case 'footer': {
            const options: any = {
              x: 0.3,
              y: 5.625 - footerBarH + 0.05,
              w: 9.4,
              h: 0.2,
              fontSize: 12,
              color: 'FFFFFF',
              ...(el.options || {})
            };
            slide.addText(el.text || '', options);
            break;
          }
        }
      });
    }
    if (slideModel.notes && slideModel.notes.length) {
      slide.addNotes(slideModel.notes.join('\n'));
    }
  });

  return pptx;
}
