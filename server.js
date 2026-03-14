const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Ensure the Resources/Data directory exists
const dataDir = path.join(__dirname, 'Resources', 'Data');
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}

const excelFilePath = path.join(dataDir, 'Client_Contact_Details.xlsx');

app.post('/save-contact', (req, res) => {
    const { name, phone, email, project } = req.body;
    const newRow = [name, phone, email, project];

    let workbook;
    let worksheet;

    // If the file already exists, read it and append the new data
    if (fs.existsSync(excelFilePath)) {
        workbook = XLSX.readFile(excelFilePath);
        worksheet = workbook.Sheets["Contact Details"];
        XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 }); // Append to the bottom
    } else {
        // If the file doesn't exist, create it with headers
        const headers = ["Your Name", "Phone Number with ISD code", "Email Address", "Tell us about your project"];
        const excelData = [headers, newRow];
        worksheet = XLSX.utils.aoa_to_sheet(excelData);
        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Contact Details");
    }

    // Save the file to the Resources/Data folder
    XLSX.writeFile(workbook, excelFilePath);

    res.status(200).json({ message: 'Success! Data saved to Resources/Data/Client_Contact_Details.xlsx' });
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});