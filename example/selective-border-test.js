/* eslint-disable no-console */
const fs = require('fs');
// FIXME: Incase you have the npm package
// const HTMLtoDOCX = require('html-to-docx');
const HTMLtoDOCX = require('../dist/html-to-docx.umd.js');

const filePath = './selective-border-test.docx';

const htmlString = `
<!DOCTYPE html>
<html>
<head>
  <title>Selective Border Test</title>
</head>
<body>
  <h1>Selective Border Direction Test</h1>
  
  <h2>Table with only top border</h2>
  <table style="border-top: 2px solid red;">
    <tr>
      <td>Cell 1</td>
      <td>Cell 2</td>
    </tr>
    <tr>
      <td>Cell 3</td>
      <td>Cell 4</td>
    </tr>
  </table>
  
  <h2>Table with only left and right borders</h2>
  <table style="border-left: 1px dashed blue; border-right: 1px dashed blue;">
    <tr>
      <td>Cell 1</td>
      <td>Cell 2</td>
    </tr>
    <tr>
      <td>Cell 3</td>
      <td>Cell 4</td>
    </tr>
  </table>
  
  <h2>Cell with only bottom border</h2>
  <table>
    <tr>
      <td style="border-bottom: 3px double green;">Cell with bottom border</td>
      <td>Normal cell</td>
    </tr>
    <tr>
      <td>Normal cell</td>
      <td>Normal cell</td>
    </tr>
  </table>
  
  <h2>Mixed cell borders</h2>
  <table>
    <tr>
      <td style="border-top: 1px solid black; border-left: 2px dotted red;">Top + Left</td>
      <td style="border-right: 1px solid blue;">Right only</td>
    </tr>
    <tr>
      <td style="border-bottom: 2px dashed orange;">Bottom only</td>
      <td style="border: none;">No border</td>
    </tr>
  </table>
  
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

    console.log('Selective border test document created successfully!');
    console.log(`File saved as: ${filePath}`);
  } catch (error) {
    console.log('Error:', error);
  }
})();