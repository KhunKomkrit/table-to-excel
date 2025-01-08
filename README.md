# HTML to Excel Converter

This project provides a utility to convert HTML tables into Excel spreadsheets, complete with styles like background colors and font colors. The generated file can be returned as a Blob, making it compatible with both Node.js and browser environments.

---

## Features
- Parses HTML tables and converts them to Excel files.
- Supports inline styles like `background-color` and `color`.
- Works in both **Node.js** and **browser** environments.
- Returns the Excel file as a `Blob`.

---

## Installation

Install the package using npm:

```bash
npm install html-to-excel-converter
```

---

## Usage

### Node.js Example

```typescript
import { htmlToExcelBlob } from 'html-to-excel-converter';
import fs from 'fs';

const html = `
<table>
  <tr>
    <th style="background-color: #ff9999; color: white;">Name</th>
    <th style="background-color: #ff9999; color: white;">Age</th>
  </tr>
  <tr>
    <td style="background-color: #ffe6e6;">John</td>
    <td style="background-color: #ffe6e6;">30</td>
  </tr>
</table>
`;

htmlToExcelBlob(html).then(async (blob) => {
  const buffer = Buffer.from(await blob.arrayBuffer());
  fs.writeFileSync('output.xlsx', buffer);
  console.log('Excel file created successfully as output.xlsx');
});
```

### Browser Example

```typescript
import { htmlToExcelBlob } from 'html-to-excel-converter';

const html = `
<table>
  <tr>
    <th style="background-color: #ff9999; color: white;">Name</th>
    <th style="background-color: #ff9999; color: white;">Age</th>
  </tr>
  <tr>
    <td style="background-color: #ffe6e6;">John</td>
    <td style="background-color: #ffe6e6;">30</td>
  </tr>
</table>
`;

htmlToExcelBlob(html).then((blob) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'output.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  console.log('Excel file downloaded successfully');
});
```

---

## API

### `htmlToExcelBlob(html: string): Promise<Blob>`
Converts an HTML string containing a `<table>` to an Excel file and returns it as a `Blob`.

#### Parameters:
- `html` (string): The HTML string containing the table to convert.

#### Returns:
- `Promise<Blob>`: A Blob representing the Excel file.

---

## Requirements
- Node.js 14 or later (for Node.js usage).
- A modern browser (for browser usage).

---

## Development

### Setup
Clone the repository and install dependencies:

```bash
git clone https://github.com/your-username/html-to-excel-converter.git
cd html-to-excel-converter
npm install
```

### Run Tests

To run the tests:

```bash
npm test
```

---

## Contribution
Contributions are welcome! Please follow these steps:
1. Fork the repository.
2. Create a new branch for your feature: `git checkout -b feature-name`.
3. Commit your changes: `git commit -m "Add feature-name"`.
4. Push to your branch: `git push origin feature-name`.
5. Create a pull request.

---

## License
This project is licensed under the [MIT License](LICENSE).

---

## Acknowledgments
- [ExcelJS](https://github.com/exceljs/exceljs) for Excel file generation.
- [JSDOM](https://github.com/jsdom/jsdom) for HTML parsing in Node.js.

---

Enjoy converting HTML tables to Excel with ease!

