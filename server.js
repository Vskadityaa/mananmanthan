const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Define the path for the Excel file
const excelFile = path.join(__dirname, 'registrations.xlsx');

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public')); // Serve static files from 'public' folder

// POST route to save form data
app.post('/register', (req, res) => {
  const { name, email, phone, address, style } = req.body;
  console.log("Received data:", req.body);

  let data = [];

  // Read existing data if file exists
  if (fs.existsSync(excelFile)) {
    const workbook = XLSX.readFile(excelFile);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet);
  }

  // Add new registration
  data.push({
    Name: name,
    Email: email,
    Phone: phone,
    Address: address,
    DanceStyle: style
  });

  // Write updated data to Excel
  const newWB = XLSX.utils.book_new();
  const newWS = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(newWB, newWS, "Registrations");
  XLSX.writeFile(newWB, excelFile);

  // Redirect to index.html after registration
  res.redirect('/index.html');
});

// Start the server
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}/`);
});
