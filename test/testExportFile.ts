import { htmlToExcel } from '../src/index';

const html = `
<table style="border: 1px solid black; background-color: #ffcccc;">
  <thead>
    <tr>
      <th style="border: 1px solid black; background-color: #ffffff; color: white;" rowspan="2">Header 1</th>
      <th style="border: 1px solid black; background-color: #ffffff; color: white;" colspan="2">Header Group 2</th>
      <th style="border: 1px solid black; background-color: #ffffff; color: white;" rowspan="2">Header 3</th>
    </tr>
    <tr>
      <th style="border: 1px solid black; background-color: #ffffff; color: white;">Subheader 2.1</th>
      <th style="border: 1px solid black; background-color: #ffffff; color: white;">Subheader 2.2</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="2">Row 1, Col 1</td>
      <td style="border: 1px solid black; background-color: #ffe6e6;">Row 1, Col 2.1</td>
      <td rowspan="2">Row 1-2, Col 2.2</td>
      <td>Row 1, Col 3</td>
    </tr>
    <tr>
      <td>Row 2, Col 2.1</td>
      <td>Row 2, Col 3</td>
    </tr>
    <tr>
      <td>Row 3, Col 1</td>
      <td colspan="2">Row 3, Col 2.1-2.2</td>
      <td>Row 3, Col 3</td>
    </tr>
    <tr>
      <td colspan="4">Footer Row (Merged Across All Columns)</td>
    </tr>
  </tbody>
</table>
`;

htmlToExcel(html, './test/output.xlsx').then(() => {
  console.log('Excel file created successfully!');
});
