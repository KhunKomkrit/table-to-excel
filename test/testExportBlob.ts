import { htmlToExcelBlob } from '../src/index';
import fs from 'fs';

const html = `
<table style="border: 1px solid black; background-color: #ffcccc;">
  <tr>
    <th style="border: 1px solid black; background-color: #ffffff; color: white;">Name</th>
    <th style="border: 1px solid black; background-color: #ffffff; color: white;">Age</th>
  </tr>
  <tr>
    <td style="border: 1px solid black; background-color: #ffe6e6;">ไทย</td>
    <td style="border: 1px solid black; background-color: #ffe6e6;">30</td>
  </tr>
</table>
`;

htmlToExcelBlob(html).then(async (blob) => {
  const buffer = Buffer.from(await blob.arrayBuffer());
  fs.writeFileSync('test/output-blob.xlsx', buffer);
  console.log('Excel file created successfully as output-blob.xlsx');
});
