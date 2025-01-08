import { htmlToExcel } from '../src/index';

const html = `
<table style="border: 1px solid black; background-color: #ffcccc;">
  <tr>
    <th style="background-color: #ffffff; color: white;">Name</th>
    <th style="background-color: #ffffff; color: white;">Age</th>
  </tr>
  <tr>
    <td style="background-color: #ffe6e6;">ไทย</td>
    <td style="background-color: #ffe6e6;">30</td>
  </tr>
</table>
`;

htmlToExcel(html, './test/output.xlsx').then(() => {
  console.log('Excel file created successfully!');
});
