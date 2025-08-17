/**
 * Build a PPTX presentation from a structured slide model.
 */

export interface SlideElement {
  type: 'title' | 'text' | 'image';
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

  slides.forEach(slideModel => {
    const slide = pptx.addSlide();
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
            const options: any = { path: el.src, x: 0.5, y, w: 4, h: 3, ...(el.options || {}) };
            slide.addImage(options);
            y += options.h + 0.5;
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
