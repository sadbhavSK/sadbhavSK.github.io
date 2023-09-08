const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const port = 80;

app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'pug');
app.set('views', path.join(__dirname, 'views'));

app.get('/', (req, res) => {
    res.status(200).render('index.pug');
});

app.post('/submitted', async (req, res) => {
    try {
        const formInput = req.body;
        console.log(formInput);

        // Load the existing Excel file if it exists, or create a new one if not
        const filePath = path.join(__dirname, 'file', 'data.xlsx');
        let workbook;

        if (fs.existsSync(filePath)) {
            workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
        } else {
            workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('FormResponses');

            // Add headers with formatting
            const headerRow = worksheet.addRow(['Name', 'Email', 'Password', 'Username', 'DOB']);
            headerRow.font = { bold: true };
            headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
            headerRow.height = 20;

            // Apply cell border formatting to headers
            headerRow.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' },
                };
            });
        }

        const worksheet = workbook.getWorksheet('FormResponses');

        // Add the form data to the Excel sheet
        const newRow = worksheet.addRow([
            formInput.name,
            formInput.email,
            formInput.password,
            formInput.username,
            formInput.dob,
        ]);

        // Apply cell border formatting to the new row
        newRow.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' },
            };
        });

        // Save the updated Excel file
        await workbook.xlsx.writeFile(filePath);

        res.render('afterSubmit');
    } catch (error) {
        console.error(error);
        res.status(500).send('An error occurred.');
    }
});

app.listen(port, () => {
    console.log(`App is successfully serving at port ${port}`);
});
