const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const app = express();
const port = 3000;

app.use(cors());
app.use(express.json()); // Middleware to parse JSON requests

// Path to the Excel file
const excelFilePath = 'C:\\Users\\dhana\\OneDrive\\Desktop\\patient-ifo-app\\ozone.xlsx';
const workbook = xlsx.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const patientData = xlsx.utils.sheet_to_json(sheet);

app.get('/get_patient_info', (req, res) => {
  const admissionNo = req.query.admissionNo;
  const patient = patientData.find(p => p.admissionNo == admissionNo);
  if (patient) {
    res.json(patient);
  } else {
    res.status(404).json({ error: 'Patient not found' });
  }
});

app.post('/update-excel', (req, res) => {
  console.log('POST request received at /update-excel');
  console.log('Request body:', req.body);
  const { admissionNo, followUpCharges, followUpChargesAfterTDS, calculationTime } = req.body;

  // Find the row index corresponding to the admission number
  const rowIndex = patientData.findIndex(p => p.admissionNo == admissionNo);

  if (rowIndex !== -1) {
    // Update the follow-up charges, follow-up charges after TDS deduction, and calculation time columns
    const rowNumber = rowIndex + 2; // Excel rows start from 1, and header is row 1
    sheet[`AG${rowNumber}`] = { v: followUpCharges }; // Assuming 'Follow-up Charges' column is AG
    sheet[`AH${rowNumber}`] = { v: followUpChargesAfterTDS }; // Assuming 'Follow-up Charges after TDS Deduction' column is AH
    sheet[`AI${rowNumber}`] = { v: calculationTime }; // Assuming 'Calculation Time' column is AI

    // Write the updated workbook to the Excel file
    xlsx.writeFile(workbook, excelFilePath);
    res.json({ message: 'Excel sheet updated successfully' });
  } else {
    res.status(404).json({ error: 'Patient with provided admission number not found' });
  }
});

// Serve index.html file
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
