"use strict";
/**
 * Build a PPTX presentation from a structured slide model.
 */
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildPptx = buildPptx;
/**
 * Parse simple markdown for **bold** segments and return a value suitable for PptxGenJS addText.
 * If no markdown is present the original string is returned.
 */
function parseBold(text) {
    var regex = /\*\*(.*?)\*\*/g;
    var lastIndex = 0;
    var parts = [];
    var match;
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
function buildPptx(slides, meta) {
    if (meta === void 0) { meta = {}; }
    var pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    if (meta.title) {
        pptx.coreProps = { title: meta.title };
    }
    // Attempt to read theme colors from CSS variables; fall back to defaults
    var brand = '#1e3a8a';
    var accent = '#f97316';
    if (typeof window !== 'undefined') {
        var css = getComputedStyle(document.documentElement);
        brand = css.getPropertyValue('--brand').trim() || brand;
        accent = css.getPropertyValue('--accent').trim() || accent;
    }
    var titleBarH = 0.094;
    var accentBarH = 0.031;
    var footerBarH = 0.3125;
    slides.forEach(function (slideModel) {
        var slide = pptx.addSlide();
        // Top title and accent bars
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: titleBarH, fill: { color: brand }, line: { color: brand } });
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: titleBarH, w: 10, h: accentBarH, fill: { color: accent }, line: { color: accent } });
        // Bottom accent and footer bars
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.625 - footerBarH - accentBarH, w: 10, h: accentBarH, fill: { color: accent }, line: { color: accent } });
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.625 - footerBarH, w: 10, h: footerBarH, fill: { color: brand }, line: { color: brand } });
        if (slideModel.src) {
            // slide provided as a full-size image (e.g., html2canvas render)
            slide.addImage({ data: slideModel.src, x: 0, y: 0, w: 10, h: 5.625 });
        }
        else if (slideModel.elements) {
            var y_1 = 0.5;
            slideModel.elements.forEach(function (el) {
                switch (el.type) {
                    case 'title': {
                        var titleText = (el.text || '').replace(/\*\*(.*?)\*\*/g, '$1');
                        var options = __assign({ x: 0.5, y: y_1, w: 9, h: 1, fontSize: 32, bold: true }, (el.options || {}));
                        slide.addText(titleText, options);
                        y_1 += options.h;
                        break;
                    }
                    case 'text': {
                        var rawText = el.text || '';
                        if (el.options && typeof el.options.y === 'number') {
                            y_1 = el.options.y;
                        }
                        var lines = rawText.split('\n').length;
                        var options = __assign({ x: 0.5, y: y_1, w: 9, h: 0.6 * lines, fontSize: 18 }, (el.options || {}));
                        var formatted = parseBold(rawText);
                        slide.addText(formatted, options);
                        y_1 += options.h;
                        break;
                    }
                    case 'image': {
                        var options = __assign({ data: el.src, x: 0.5, y: y_1, w: 4, h: 3 }, (el.options || {}));
                        slide.addImage(options);
                        y_1 += options.h + 0.5;
                        break;
                    }
                    case 'footer': {
                        var options = __assign({ x: 0.3, y: 5.625 - footerBarH + 0.05, w: 9.4, h: 0.2, fontSize: 12, color: 'FFFFFF' }, (el.options || {}));
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
