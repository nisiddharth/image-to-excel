const Jimp = require('jimp');
const ExcelJS = require('exceljs');

// Function to convert RGB to hex color code
function rgbToHex(r, g, b) {
    const componentToHex = (c) => {
        const hex = c.toString(16);
        return hex.length === 1 ? '0' + hex : hex;
    };
    return '#' + componentToHex(r) + componentToHex(g) + componentToHex(b);
}

async function imageToExcel(imageFile, outputWidth, outputHeight, workbook) {
    await Jimp.read(String(imageFile)).then(image => {
        console.log("Image read")
        const originalWidth = image.bitmap.width;
        const originalHeight = image.bitmap.height;

        const widthRatio = originalWidth / outputWidth;
        const heightRatio = originalHeight / outputHeight;

        const sheet = workbook.addWorksheet('Sheet 1');
        sheet.properties.defaultRowHeight = 15;
        sheet.properties.defaultColWidth = 3;

        for (let y = 0; y < outputHeight; y++) {
            for (let x = 0; x < outputWidth; x++) {
                const startX = Math.floor(x * widthRatio);
                const endX = Math.floor((x + 1) * widthRatio);
                const startY = Math.floor(y * heightRatio);
                const endY = Math.floor((y + 1) * heightRatio);

                let totalR = 0;
                let totalG = 0;
                let totalB = 0;

                for (let px = startX; px < endX; px++) {
                    for (let py = startY; py < endY; py++) {
                        const { r, g, b } = Jimp.intToRGBA(image.getPixelColor(px, py));
                        totalR += r;
                        totalG += g;
                        totalB += b;
                    }
                }

                const averageR = Math.floor(totalR / ((endX - startX) * (endY - startY)));
                const averageG = Math.floor(totalG / ((endX - startX) * (endY - startY)));
                const averageB = Math.floor(totalB / ((endX - startX) * (endY - startY)));

                const hexColor = rgbToHex(averageR, averageG, averageB);

                const cell = sheet.getCell(y + 1, x + 1);
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: hexColor.substring(1) },
                };
            }
        }
        console.log("Image processed");
    });
}

// Add event listeners to the drag-drop area for file dragging
const dragDropArea = document.getElementById('drag-drop-area');
const fileInputDisplay = document.getElementById('file-input-display');

dragDropArea.addEventListener('dragenter', (event) => {
    event.preventDefault();
    dragDropArea.classList.add('active');
});

dragDropArea.addEventListener('dragover', (event) => {
    event.preventDefault();
    dragDropArea.classList.add('active');
});

dragDropArea.addEventListener('dragleave', (event) => {
    event.preventDefault();
    dragDropArea.classList.remove('active');
});

dragDropArea.addEventListener('drop', (event) => {
    event.preventDefault();
    dragDropArea.classList.remove('active');

    const files = event.dataTransfer.files;
    if (files.length > 0) {
        const imageFile = files[0];
        if (imageFile.type.startsWith('image/')) {
            document.getElementById('file-input').files = files;
            fileInputDisplay.textContent = imageFile.name;
        } else {
            fileInputDisplay.textContent = "No file selected";
            alert('Only image files are supported.');
        }
    }
});

document.getElementById('file-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        fileInputDisplay.textContent = file.name;
    } else {
        fileInputDisplay.textContent = 'No file selected';
    }
});



// Convert image to Excel on Convert button click
document.getElementById('convert-button').addEventListener('click', async () => {
    const outputWidth = parseInt(document.getElementById('output-width').value);
    const outputHeight = parseInt(document.getElementById('output-height').value);

    const imageInput = document.getElementById('file-input');
    if (imageInput.files.length === 0) {
        alert('Please select an image file.');
        return;
    }

    var tmppath = URL.createObjectURL(imageInput.files[0]);

    console.log("init" + tmppath);
    const workbook = new ExcelJS.Workbook();
    await imageToExcel(tmppath, outputWidth, outputHeight, workbook).then(() => {
        workbook.xlsx.writeBuffer()
            .then(buffer => {
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = 'output.xlsx';
                link.click();
            })
            .catch(err => console.log('Error exporting Excel workbook.', err));
    }).catch(err => {
        console.error(err);
        alert("Could not process the image. Supports only PNG, JPEG, BMP, and TIFF images.");
    });
});
