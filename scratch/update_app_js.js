const fs = require('fs');

const js = fs.readFileSync('js/app.js', 'utf-8');
const b64 = fs.readFileSync('resources/logo-epe-blanco.jpg').toString('base64');
const imgData = 'data:image/jpeg;base64,' + b64;

const targetMarker = '    function drawEpeLogo(doc, boxX, boxY, boxW) {';
const targetEnd = '      doc.setTextColor(0, 0, 0);\n    }';

const startIdx = js.indexOf(targetMarker);
if (startIdx === -1) {
  console.error('Start marker not found');
  process.exit(1);
}
const endIdx = js.indexOf(targetEnd, startIdx);
if (endIdx === -1) {
  console.error('End marker not found');
  process.exit(1);
}

const before = js.substring(0, startIdx);
const after = js.substring(endIdx + targetEnd.length);

const replacement = `const LOGO_EPE_BLANCO_DATA = '${imgData}';\n\n` +
`    function drawEpeLogo(doc, boxX, boxY, boxW, boxH = 19) {
      try {
        const maxImgW = 26;
        const maxImgH = 15;
        const finalH = Math.min(maxImgW * (175 / 258), maxImgH);
        const finalW = finalH * (258 / 175);
        const imgX = boxX + (boxW - finalW) / 2;
        const imgY = boxY + (boxH - finalH) / 2;
        doc.addImage(LOGO_EPE_BLANCO_DATA, 'JPEG', imgX, imgY, finalW, finalH);
      } catch (err) {
        const centerX = boxX + boxW / 2;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(16);
        doc.setTextColor(0, 75, 135);
        doc.text('EPE', centerX, boxY + 8.5, { align: 'center' });
        doc.setFillColor(0, 75, 135);
        doc.rect(boxX + 4, boxY + 10.2, boxW - 11, 1.3, 'F');
        doc.triangle(boxX + boxW - 7, boxY + 9.7, boxX + boxW - 3.5, boxY + 10.85, boxX + boxW - 7, boxY + 12.0, 'F');
        doc.setFontSize(6.5);
        doc.text('Energía de Santa Fe', centerX, boxY + 15.8, { align: 'center' });
        doc.setTextColor(0, 0, 0);
      }
    }`;

fs.writeFileSync('js/app.js', before + replacement + after, 'utf-8');
console.log('Successfully injected LOGO_EPE_BLANCO_DATA into js/app.js!');

