const express = require('express');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');
const bodyParser = require('body-parser');

const app = express();
app.use(cors());
app.use(bodyParser.json());

const FILE = 'attendance.xlsx';

function appendToExcel(data) {
  let workbook;
  let worksheet;

  if (fs.existsSync(FILE)) {
    workbook = XLSX.readFile(FILE);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Attendance');
  }

  const oldData = XLSX.utils.sheet_to_json(worksheet);
  oldData.push(data);
  const newSheet = XLSX.utils.json_to_sheet(oldData);
  workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  XLSX.writeFile(workbook, FILE);
}

app.post('/submit', (req, res) => {
  console.log("📩 Received:", req.body);
  appendToExcel(req.body);
  res.json({ message: 'Saved to Excel ✅' });
});

app.listen(3000, () => {
  console.log('✅ Server running on http://localhost:3000');
});
