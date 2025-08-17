/**
 * Build a PPTX presentation from a structured slide model.
 */

export interface SlideElement {
  type: 'title' | 'text' | 'image';
  text?: string;
  src?: string;
  options?: any;
}

export interface SlideModel {
  elements: SlideElement[];
  notes?: string[];
}

// PptxGenJS is loaded globally via <script>
declare const PptxGenJS: any;

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
    let y = 0.5;
    slideModel.elements.forEach(el => {
      switch (el.type) {
        case 'title':
          slide.addText(el.text || '', { x: 0.5, y, w: 9, fontSize: 32, bold: true, ...(el.options || {}) });
          y += 1;
          break;
        case 'text':
          slide.addText(el.text || '', { x: 0.5, y, w: 9, fontSize: 18, ...(el.options || {}) });
          y += 0.6;
          break;
        case 'image':
          slide.addImage({ path: el.src, x: 0.5, y, w: 4, h: 3, ...(el.options || {}) });
          y += 3.5;
          break;
      }
    });
    if (slideModel.notes && slideModel.notes.length) {
      slide.addNotes(slideModel.notes.join('\n'));
    }
  });

  return pptx;
}
