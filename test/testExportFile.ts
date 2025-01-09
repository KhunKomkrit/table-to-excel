import { htmlToExcel } from '../src/index';

const html = `
<table border="1" cellspacing="0" cellpadding="5">
  <thead>
    <tr>
      <th rowspan="2">Header 1</th>
      <th colspan="2">Header Group 2</th>
      <th rowspan="2">Header 3</th>
    </tr>
    <tr>
      <th>Subheader 2.1</th>
      <th>Subheader 2.2</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="2">Row 1, Col 1</td>
      <td>Row 1, Col 2.1</td>
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
