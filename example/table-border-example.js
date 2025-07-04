/* eslint-disable no-console */
const fs = require('fs');
// FIXME: Incase you have the npm package
// const HTMLtoDOCX = require('html-to-docx');
const HTMLtoDOCX = require('../dist/html-to-docx.umd.js');

const filePath = './table-border-example.docx';

const htmlString = `<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <title>Table Border Style Test</title>
    </head>
    <body>
        <h1>Table Border Style Test</h1>
        
        <h2>1. Basic Table with Solid Borders</h2>
        <table style="border: 2px solid #000000; border-collapse: collapse;">
            <tr>
                <th style="border: 1px solid #000000; padding: 8px;">Header 1</th>
                <th style="border: 1px solid #000000; padding: 8px;">Header 2</th>
            </tr>
            <tr>
                <td style="border: 1px solid #000000; padding: 8px;">Cell 1</td>
                <td style="border: 1px solid #000000; padding: 8px;">Cell 2</td>
            </tr>
        </table>
        
        <h2>2. Table with Dashed Borders</h2>
        <table style="border: 2px dashed #ff0000; border-collapse: collapse;">
            <tr>
                <th style="border: 1px dashed #ff0000; padding: 8px;">Header 1</th>
                <th style="border: 1px dashed #ff0000; padding: 8px;">Header 2</th>
            </tr>
            <tr>
                <td style="border: 1px dashed #ff0000; padding: 8px;">Cell 1</td>
                <td style="border: 1px dashed #ff0000; padding: 8px;">Cell 2</td>
            </tr>
        </table>
        
        <h2>3. Table with Dotted Borders</h2>
        <table style="border: 3px dotted #0000ff; border-collapse: collapse;">
            <tr>
                <th style="border: 2px dotted #0000ff; padding: 8px;">Header 1</th>
                <th style="border: 2px dotted #0000ff; padding: 8px;">Header 2</th>
            </tr>
            <tr>
                <td style="border: 2px dotted #0000ff; padding: 8px;">Cell 1</td>
                <td style="border: 2px dotted #0000ff; padding: 8px;">Cell 2</td>
            </tr>
        </table>
        
        <h2>4. Table with Double Borders</h2>
        <table style="border: 4px double #008000; border-collapse: collapse;">
            <tr>
                <th style="border: 3px double #008000; padding: 8px;">Header 1</th>
                <th style="border: 3px double #008000; padding: 8px;">Header 2</th>
            </tr>
            <tr>
                <td style="border: 3px double #008000; padding: 8px;">Cell 1</td>
                <td style="border: 3px double #008000; padding: 8px;">Cell 2</td>
            </tr>
        </table>
        
        <h2>5. Table with Mixed Border Styles</h2>
        <table style="border-collapse: collapse;">
            <tr>
                <th style="border-top: 3px solid #000000; border-left: 2px dashed #ff0000; border-right: 2px dotted #0000ff; border-bottom: 1px solid #000000; padding: 8px;">Mixed Header 1</th>
                <th style="border-top: 3px solid #000000; border-left: 2px dotted #0000ff; border-right: 2px dashed #ff0000; border-bottom: 1px solid #000000; padding: 8px;">Mixed Header 2</th>
            </tr>
            <tr>
                <td style="border-top: 1px solid #000000; border-left: 2px dashed #ff0000; border-right: 2px dotted #0000ff; border-bottom: 2px double #008000; padding: 8px;">Mixed Cell 1</td>
                <td style="border-top: 1px solid #000000; border-left: 2px dotted #0000ff; border-right: 2px dashed #ff0000; border-bottom: 2px double #008000; padding: 8px;">Mixed Cell 2</td>
            </tr>
        </table>
        
        <h2>6. Table with No Borders (None)</h2>
        <table style="border: none; border-collapse: collapse;">
            <tr>
                <th style="border: none; padding: 8px; background-color: #f0f0f0;">No Border Header 1</th>
                <th style="border: none; padding: 8px; background-color: #f0f0f0;">No Border Header 2</th>
            </tr>
            <tr>
                <td style="border: none; padding: 8px;">No Border Cell 1</td>
                <td style="border: none; padding: 8px;">No Border Cell 2</td>
            </tr>
        </table>
        
        <h2>7. Table with Selective Border Removal</h2>
        <table style="border: 2px solid #000000; border-collapse: collapse;">
            <tr>
                <th style="border: 2px solid #000000; border-bottom: none; padding: 8px;">Header 1 (No Bottom)</th>
                <th style="border: 2px solid #000000; border-left: none; padding: 8px;">Header 2 (No Left)</th>
            </tr>
            <tr>
                <td style="border: 2px solid #000000; border-top: none; border-right: none; padding: 8px;">Cell 1 (No Top/Right)</td>
                <td style="border: 2px solid #000000; border-bottom: none; border-left: none; padding: 8px;">Cell 2 (No Bottom/Left)</td>
            </tr>
        </table>
        
        <h2>8. Table with Different Border Widths</h2>
        <table style="border-collapse: collapse;">
            <tr>
                <th style="border-top: 5px solid #000000; border-left: 3px solid #000000; border-right: 3px solid #000000; border-bottom: 1px solid #000000; padding: 8px;">Thick Top Border</th>
                <th style="border-top: 5px solid #000000; border-left: 3px solid #000000; border-right: 3px solid #000000; border-bottom: 1px solid #000000; padding: 8px;">Thick Top Border</th>
            </tr>
            <tr>
                <td style="border-top: 1px solid #000000; border-left: 3px solid #000000; border-right: 3px solid #000000; border-bottom: 5px solid #000000; padding: 8px;">Thick Bottom Border</td>
                <td style="border-top: 1px solid #000000; border-left: 3px solid #000000; border-right: 3px solid #000000; border-bottom: 5px solid #000000; padding: 8px;">Thick Bottom Border</td>
            </tr>
        </table>
    </body>
</html>`;

(async () => {
  try {
    const docxBuffer = await HTMLtoDOCX(htmlString, null, {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: true,
    });

    fs.writeFileSync(filePath, docxBuffer);

    console.log('Table border test document created successfully!');
    console.log(`File saved as: ${filePath}`);
  } catch (error) {
    console.error('Error creating document:', error);
  }
})();