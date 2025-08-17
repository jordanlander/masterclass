/**
 * Build an ODP presentation from slide images.
 */
export async function buildOdp(slides, meta = {}) {
  const zip = new JSZip();
  // The mimetype file must be stored with no compression
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation', { compression: 'STORE' });

  const manifest = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">',
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>',
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
  ];

  const content = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">',
    '<office:body><office:presentation>'
  ];

  slides.forEach((s, i) => {
    const idx = i + 1;
    const imgName = `Pictures/slide${idx}.png`;
    manifest.push(`<manifest:file-entry manifest:full-path="${imgName}" manifest:media-type="image/png"/>`);
    content.push(`<draw:page draw:name="page${idx}"><draw:frame svg:x="0cm" svg:y="0cm" svg:width="28cm" svg:height="21cm"><draw:image xlink:href="${imgName}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></draw:page>`);
    const base64 = s.src.split(',')[1];
    zip.file(imgName, base64, { base64: true });
  });

  content.push('</office:presentation></office:body></office:document-content>');
  manifest.push('</manifest:manifest>');

  zip.file('content.xml', content.join(''));
  zip.file('META-INF/manifest.xml', manifest.join(''));

  return zip.generateAsync({ type: 'blob' });
}
