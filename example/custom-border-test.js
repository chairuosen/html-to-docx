/* eslint-disable no-console */
const fs = require('fs');
// FIXME: Incase you have the npm package
// const HTMLtoDOCX = require('html-to-docx');
const HTMLtoDOCX = require('../dist/html-to-docx.umd.js');

const filePath = './custom-border-test.docx';

const htmlString = `
<!DOCTYPE html>
<html>
<head>
  <title>Custom Border Test</title>
</head>
<body>
  <h1>Custom Border Test - Inside H and V Borders</h1>
  
  <h2>Table with custom inside borders</h2>
  <table style="--docx-inside-h-border: 2px solid red; --docx-inside-v-border: 1px dashed blue; border-collapse: collapse;">
    <tr>
      <td>Cell 1,1</td>
      <td>Cell 1,2</td>
      <td>Cell 1,3</td>
    </tr>
    <tr>
      <td>Cell 2,1</td>
      <td>Cell 2,2</td>
      <td>Cell 2,3</td>
    </tr>
    <tr>
      <td>Cell 3,1</td>
      <td>Cell 3,2</td>
      <td>Cell 3,3</td>
    </tr>
  </table>
  
  <h2>Cell with custom inside borders</h2>
  <table>
    <tr>
      <td style="--docx-inside-h-border: 3px double green; --docx-inside-v-border: 2px solid orange;">Cell with custom inside borders</td>
      <td>Normal cell</td>
    </tr>
    <tr>
      <td>Normal cell</td>
      <td>Normal cell</td>
    </tr>
  </table>
  
  <h2>Mixed borders test</h2>
  <table style="border: 1px solid black; --docx-inside-h-border: 2px solid red;">
    <tr>
      <td style="border-top: 3px solid purple;">Custom top border</td>
      <td style="--docx-inside-v-border: 2px dashed cyan;">Custom inside V border</td>
    </tr>
    <tr>
      <td>Normal cell</td>
      <td>Normal cell</td>
    </tr>
  </table>
  
  <h2>Expected Results:</h2>
  <ul>
    <li>First table: Should have red horizontal inside borders (2px solid) and blue vertical inside borders (1px dashed)</li>
    <li>Second table: First cell should have green horizontal inside borders (3px double) and orange vertical inside borders (2px solid)</li>
    <li>Third table: Should have black outer borders, red horizontal inside borders, and the second cell should have cyan vertical inside borders</li>
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

    console.log('Custom border test document created successfully!');
    console.log(`File saved as: ${filePath}`);
  } catch (error) {
    console.log('Error:', error);
  }
})();