const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { parse } = require('csv-parse');

const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.use(express.static('public'));

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, `${Date.now()}-${file.originalname}`);
    }
});

const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        const filePath = req.file.path;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        const worksheet = workbook.getWorksheet(1);
        const csvData = [];
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber >= 4) {
                const rowData = [
                    `İstanbul Anadolu, ${row.getCell(1).value},  Swap, TE, ${row.getCell(2).value}`,
                    `İstanbul Anadolu, ${row.getCell(1).value},  Container, ET, ${row.getCell(3).value}`,
                    `İstanbul Anadolu, ${row.getCell(1).value},  Swap, TE, ${row.getCell(4).value}`,
                    `İstanbul Anadolu, ${row.getCell(1).value},  Container, ET, ${row.getCell(5).value}`,
                    `İstanbul Avrupa/Tekirdağ/Çorlu, ${row.getCell(1).value},  Swap, TE, ${row.getCell(6).value}`,
                    `İstanbul Avrupa/Tekirdağ/Çorlu, ${row.getCell(1).value},  Container, ET, ${row.getCell(7).value}`,
                    `İstanbul Avrupa/Tekirdağ/Çorlu, ${row.getCell(1).value},  Swap, TE, ${row.getCell(8).value}`,
                    `İstanbul Avrupa/Tekirdağ/Çorlu, ${row.getCell(1).value},  Container, ET, ${row.getCell(9).value}`,
                    `İzmir , ${row.getCell(1).value},  Swap, TE, ${row.getCell(10).value}`,
                    `İzmir , ${row.getCell(1).value},  Container, ET, ${row.getCell(11).value}`,
                    `İzmir , ${row.getCell(1).value},  Swap, TE, ${row.getCell(12).value}`,
                    `İzmir , ${row.getCell(1).value},  Container, ET, ${row.getCell(13).value}`,
                    `Mersin-Adana , ${row.getCell(1).value},  Swap, TE, ${row.getCell(14).value}`,
                    `Mersin-Adana , ${row.getCell(1).value},  Container, ET, ${row.getCell(15).value}`,
                    `Mersin-Adana , ${row.getCell(1).value},  Swap, TE, ${row.getCell(16).value}`,
                    `Mersin-Adana , ${row.getCell(1).value},  Container, ET, ${row.getCell(17).value}`
                ];
                csvData.push(rowData.join('\n'));
            }
        });

        // Manuel başlık
        const headers = ['Bölge', 'Zip Code', 'Araç tipi', 'Taşıma', 'Tutar']; // Bu başlıkları manuel olarak belirleyin
        const csvContent = [headers.join(',')].concat(csvData).join('\n');
        const csvFilePath = path.join(__dirname, 'uploads', `${Date.now()}-output.csv`);
        fs.writeFileSync(csvFilePath, csvContent);

        const csvFileContent = fs.readFileSync(csvFilePath, 'utf8');

        parse(csvFileContent, { columns: true, trim: true }, (err, output) => {
            if (err) {
                return res.status(500).json({ error: 'CSV parsing error' });
            }

            res.render('result', { csvFileName: path.basename(csvFilePath), csvData: output });
        });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'File processing error' });
    }
});

app.get('/uploads/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    res.download(filePath);
});

app.listen(port, () => {
    console.log(`Sunucu ${port} portunda çalışıyor.`);
});
