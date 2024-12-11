const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Serve static files (e.g., CSS, images, and JS)
app.use(express.static(path.join(__dirname, 'public')));

// Serve the contact form
app.get('/contact.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'contact.html'));
});

// Handle form submission
app.post('/submit-contact', (req, res) => {
    const { name, email, message } = req.body;

    // Create or append to an Excel file
    const filePath = path.join(__dirname, 'contacts.xlsx');
    let workbook, worksheet;

    // Check if the file already exists
    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
        worksheet = workbook.Sheets['Contacts'];
    } else {
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Contacts');
    }

    // Append the new contact
    const data = XLSX.utils.sheet_to_json(worksheet);
    data.push({ Name: name, Email: email, Message: message });
    const newSheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets['Contacts'] = newSheet;

    // Save the updated file
    XLSX.writeFile(workbook, filePath);

    res.send('Thank you for contacting us!');
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
