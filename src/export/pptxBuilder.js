/**
 * Build a PPTX presentation from a structured slide model.
 */
function parseBold(text) {
  const regex = /\*\*(.*?)\*\*/g;
  let lastIndex = 0;
  const parts = [];
  let match;
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
          case 'title': {
            const titleText = (el.text || '').replace(/\*\*(.*?)\*\*/g, '$1');
            const options = { x: 0.5, y, w: 9, h: 1, fontSize: 32, bold: true, ...(el.options || {}) };
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
            const options = { x: 0.5, y, w: 9, h: 0.6 * lines, fontSize: 18, ...(el.options || {}) };
            const formatted = parseBold(rawText);
            slide.addText(formatted, options);
            y += options.h;
            break;
          }
          case 'image': {
            const options = { path: el.src, x: 0.5, y, w: 4, h: 3, ...(el.options || {}) };
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
