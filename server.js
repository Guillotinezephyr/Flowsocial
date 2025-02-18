const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));

app.post('/submit-form', (req, res) => {
  const { firstName, lastName, email, phoneNumber } = req.body;

  // Create an Excel file
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet([{ firstName, lastName, email, phoneNumber }]);
  xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
  const filePath = 'form_data.xlsx';
  xlsx.writeFile(wb, filePath);

  res.send('Form data saved to Excel file.');
});

app.listen(3000, () => {
  console.log('Server is running on port 3000');
});