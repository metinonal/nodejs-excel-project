// public/scripts.js

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('excelFile').addEventListener('change', handleFileSelect, false);
});

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = new ExcelJS.Workbook();
            workbook.xlsx.load(data).then(workbook => {
                const worksheet = workbook.getWorksheet(1);
                const sheetData = [];
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
                        sheetData.push(rowData.join('\n'));
                    }
                });
                displayPreview(sheetData);
                hideLoading();
            });
        };
        showLoading();
        reader.readAsArrayBuffer(file);
    }
}

function displayPreview(data) {
    const table = document.getElementById('filePreview');
    table.innerHTML = '';
    data.forEach(row => {
        const rowElement = document.createElement('tr');
        row.split('\n').forEach(cell => {
            const cellElement = document.createElement('td');
            cellElement.textContent = cell;
            rowElement.appendChild(cellElement);
        });
        table.appendChild(rowElement);
    });
}

function showLoading() {
    document.getElementById('loading').style.display = 'block';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}
