const QRCode = require('qrcode');
const jsQR = require('jsqr');
const fs = require('fs').promises;
const Jimp = require('jimp');
const ExcelJS = require('exceljs');

// Function to generate QR code and print to the console with a smaller size
async function generateAndPrintQRCode(data, size = 3, filePath) {
    try {
        const qrCode = await QRCode.toFile(filePath, data, { scale: size });
        console.log('QR code saved to:', filePath);
        return qrCode;
    } catch (error) {
        console.error('Error generating QR code:', error.message);
        return null;
    }
}

// Function to read QR code from an image
async function readQRCode(imagePath) {
    try {
        const image = await Jimp.read(imagePath);
        const width = image.bitmap.width;
        const height = image.bitmap.height;
        const imageData = image.bitmap.data;

        const code = jsQR(imageData, width, height);

        if (code) {
            return code.data;
        } else {
            console.error('No QR code found in the image');
            return null;
        }
    } catch (error) {
        console.error('Error reading QR code:', error.message);
        return null;
    }
}

// Function to store ID and name in an Excel file
async function storeDataToExcel(data, filePath) {
    console.log('Value of data is ' + typeof(data));
    const workbook = new ExcelJS.Workbook();

    try {
        // Try loading existing workbook
        await workbook.xlsx.readFile(filePath);
    } catch (error) {
        // Workbook does not exist, create a new one
        console.log('Creating a new Excel workbook.');
    }

    let sheet = workbook.getWorksheet('Sheet1');
    if (!sheet) {
        // If the sheet doesn't exist, create a new one
        sheet = workbook.addWorksheet('Sheet1');
        sheet.addRow(['ID', 'Name']); // Add headers
    }

    // Add data
    const obj = JSON.parse(data);
    const id = obj.id;
    const name = obj.name;
    sheet.addRow([id, name]);

    // Write to file
    await workbook.xlsx.writeFile(filePath);

    console.log('Data stored in Excel file successfully.');
}

// Main function
async function main() {
    // Example data for QR code
    const exampleData = '{"id": 4, "name": "Blah Blah"}';

    // Specify the path where you want to save the QR code image
    const imagePath = `qrcode_${Date.now()}.png`;
    
    // Generate and save QR code to an image file
    await generateAndPrintQRCode(exampleData, 2, imagePath);

    // Simulate scanning the saved QR code
    const scannedData = await readQRCode(imagePath);

    // Specify the path where you want to save the Excel file
    const excelFilePath = 'data.xlsx';

    // Store scanned data in an Excel file
    if (scannedData) {
        console.log('Scanned data:', scannedData);
        await storeDataToExcel(scannedData, excelFilePath);
    }
}

// Run the main function
main();
