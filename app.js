const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.use(express.static('public'));
app.use('/uploads', express.static('uploads')); // Statik dosyalar için ekleme

// Multer ayarları
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});
const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        const filePath = req.file.path;
        const originalName = req.file.originalname;
        const csvFileName = originalName.replace(path.extname(originalName), '.csv');

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        const worksheet = workbook.getWorksheet(1);

        const csvData = [];

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber >= 4) { // 4. satırdan itibaren verileri okuyun
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

        const csvFilePath = path.join(__dirname, 'uploads', csvFileName);
        fs.writeFileSync(csvFilePath, csvData.join('\n'));

        res.render('result', { csvFileName });
    } catch (error) {
        res.status(500).send('Excel dosyası okunurken veya CSV dosyası oluşturulurken bir hata oluştu: ' + error.message);
    }
});

app.listen(port, () => {
    console.log(`Sunucu http://localhost:${port} adresinde çalışıyor.`);
});
