/* eslint-disable no-console */
const fs = require('fs');
// FIXME: Incase you have the npm package
// const HTMLtoDOCX = require('html-to-docx');
const HTMLtoDOCX = require('../dist/html-to-docx.umd.js');

const filePath = './border-fix-test.docx';

const htmlString = `
<!DOCTYPE html>
<html>
<head>
  <title>Border Fix Test</title>
</head>
<body>
  <h1>Border Fix Test - Individual Border Directions</h1>
  
  <table>
    <tr>
      <td style="border-top: 1.50pt solid #000000; border-left: none;">Top border only, left none</td>
      <td style="border-right: 2px solid red; border-bottom: none;">Right border only, bottom none</td>
    </tr>
    <tr>
      <td style="border-left: 1px dashed blue; border-top: none;">Left border only, top none</td>
      <td style="border-bottom: 3px double green; border-right: none;">Bottom border only, right none</td>
    </tr>
  </table>
  
  <h2>Expected Results:</h2>
  <ul>
    <li>First cell: Only top border should be visible (1.5pt solid black), left should be nil</li>
    <li>Second cell: Only right border should be visible (2px solid red), bottom should be nil</li>
    <li>Third cell: Only left border should be visible (1px dashed blue), top should be nil</li>
    <li>Fourth cell: Only bottom border should be visible (3px double green), right should be nil</li>
  </ul>
</body>
</html>
`;

(async () => {
  try {
    const docxBuffer = await HTMLtoDOCX(htmlString, null, {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: true,
    });

    fs.writeFileSync(filePath, docxBuffer);

    console.log('Border fix test document created successfully!');
    console.log(`File saved as: ${filePath}`);
  } catch (error) {
    console.log('Error:', error);
  }
})();