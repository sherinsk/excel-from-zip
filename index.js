const express = require('express');
const multer = require('multer');
const unzipper = require('unzipper');
const stream = require('stream');
const xlsx = require('xlsx');

const app = express();
const upload = multer();
const PORT = 3000;

// Set EJS as the templating engine
app.set('view engine', 'ejs');

// Render the upload page
app.get('/', (req, res) => {
    res.render('upload');
});

app.post('/upload', upload.single('zipfile'), (req, res) => {
    const zipBuffer = req.file.buffer;
    const readStream = new stream.PassThrough();

    readStream.end(zipBuffer);

    let jsonData = [];

    readStream
        .pipe(unzipper.Parse())
        .on('entry', (entry) => {
            const fileName = entry.path;
            if (fileName === 'students.json') {
                let data = '';

                entry.on('data', (chunk) => {
                    data += chunk;
                });

                entry.on('end', () => {
                    jsonData = JSON.parse(data);
                    convertJsonToExcel(jsonData, res);
                });
            } else {
                entry.autodrain();
            }
        })
        .on('error', (err) => {
            console.error(err);
            res.status(500).send('Error processing the ZIP file');
        });
});

// Helper function to safely get the length of a value
const getValueLength = (value) => {
    if (value == null) return 0; // Handle null/undefined
    if (typeof value === 'object') return JSON.stringify(value).length; // Handle objects/arrays
    return value.toString().length; // Handle strings and numbers
};

const convertJsonToExcel = (jsonData, res) => {
    const worksheet = xlsx.utils.json_to_sheet(jsonData);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Students');

    // Set column widths with proper type checking and optimized calculation
    const colWidths = Object.keys(jsonData[0] || {}).map((key) => {
        try {
            const maxLength = Math.max(
                key.length, // Header length
                ...jsonData.slice(0, 1000).map(row => getValueLength(row[key])) // Limit to first 1000 rows to prevent large data from causing stack overflow
            );
            return { wch: maxLength + 2 }; // Add extra space for readability
        } catch (error) {
            console.error('Error calculating column width:', error);
            return { wch: 10 }; // Default width in case of error
        }
    });
    worksheet['!cols'] = colWidths;

    const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Disposition', 'attachment; filename=students.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(excelBuffer);
};

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
