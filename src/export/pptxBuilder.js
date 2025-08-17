/**
 * Build a PPTX presentation from a structured slide model.
 */
export function buildPptx(slides, meta = {}) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  if (meta.title) {
    pptx.coreProps = { title: meta.title };
  }

  slides.forEach(slideModel => {
    const slide = pptx.addSlide();
    if (slideModel.src) {
      // Slide represented as a pre-rendered image
      slide.addImage({ data: slideModel.src, x: 0, y: 0, w: 10, h: 5.625 });
    } else if (slideModel.elements) {
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
    }
    if (slideModel.notes && slideModel.notes.length) {
      slide.addNotes(slideModel.notes.join('\n'));
    }
  });

  return pptx;
}
