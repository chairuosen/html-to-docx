const HTMLtoDOCX = require('./dist/html-to-docx.umd.js');
const fs = require('fs');

// Test the specific case that's causing issues
const htmlString = `
<table style="--docx-inside-h-border: 0.50pt solid #000000;">
  <tr>
    <td>Cell 1</td>
    <td>Cell 2</td>
  </tr>
  <tr>
    <td>Cell 3</td>
    <td>Cell 4</td>
  </tr>
</table>
`;

console.log('Testing border conversion for: --docx-inside-h-border: 0.50pt solid #000000');

(async () => {
  try {
    const docxBuffer = await HTMLtoDOCX(htmlString);
    fs.writeFileSync('debug-border-test.docx', docxBuffer);
    console.log('Generated debug-border-test.docx');
  } catch (error) {
    console.log('Error:', error);
  }
})();

// Let's also test the unit conversion directly
// Manual calculation: 1pt = 8 EIP
console.log('0.50pt should convert to EIP:', 0.50 * 8, '(expected: 4)');
console.log('1.00pt should convert to EIP:', 1.00 * 8, '(expected: 8)');
console.log('2.00pt should convert to EIP:', 2.00 * 8, '(expected: 16)');