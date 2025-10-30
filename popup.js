let currentCsvData = null;
let currentDelimiter = null;
let fontControlsLocked = false;

const tableContainer = document.getElementById('tableContainer');
const exportButtons = document.getElementById('exportButtons');
const fontFamilySelect = document.getElementById('fontFamily');
const fontSizeSelect = document.getElementById('fontSize');
const columnAlignmentInput = document.getElementById('columnAlignment');
const tableCaptionInput = document.getElementById('tableCaption');
const tableNoteTextarea = document.getElementById('tableNote');
const darkModeToggle = document.getElementById('darkModeToggle');
const textColorInput = document.getElementById('textColor');
const borderColorInput = document.getElementById('borderColor'); // NEW: Border color input

document.getElementById('dataFile').addEventListener('change', handleFileUpload);
document.getElementById('processRaw').addEventListener('click', () => {
    handleFileUpload(null);
});
document.querySelectorAll('input[name="inputMode"]').forEach(radio => {
    radio.addEventListener('change', switchInputMode);
});

document.getElementById('exportWord').addEventListener('click', exportToWord);
document.getElementById('exportPng').addEventListener('click', exportToPng);
document.getElementById('copyPng').addEventListener('click', copyPngToClipboard);

columnAlignmentInput.addEventListener('input', redrawTable);
tableCaptionInput.addEventListener('input', redrawTable);
tableNoteTextarea.addEventListener('input', redrawTable);
textColorInput.addEventListener('input', redrawTable);
borderColorInput.addEventListener('input', redrawTable); // NEW: Redraw on border color change

fontFamilySelect.addEventListener('change', () => {
    if (!fontControlsLocked) {
        redrawTable();
    }
});
fontSizeSelect.addEventListener('change', () => {
    if (!fontControlsLocked) {
        redrawTable();
    }
});

darkModeToggle.addEventListener('change', toggleDarkMode);

function toggleDarkMode() {
    document.body.classList.toggle('dark-mode', darkModeToggle.checked);
    localStorage.setItem('darkMode', darkModeToggle.checked);
    redrawTable();
}

function loadDarkModePreference() {
    const isDarkMode = localStorage.getItem('darkMode') === 'true';
    darkModeToggle.checked = isDarkMode;
    document.body.classList.toggle('dark-mode', isDarkMode);
}
loadDarkModePreference();

function setFontControlsLock(lock) {
    fontControlsLocked = lock;
    fontFamilySelect.disabled = lock;
    fontSizeSelect.disabled = lock;
}


function switchInputMode(event) {
    const mode = event.target.value;
    document.getElementById('fileMode').style.display = mode === 'file' ? 'block' : 'none';
    document.getElementById('rawMode').style.display = mode === 'raw' ? 'block' : 'none';
}

function handleFileUpload(event) {
    const inputMode = document.querySelector('input[name="inputMode"]:checked').value;
    let dataPromise;

    if (inputMode === 'file' && event && event.target.files.length > 0) {
        const file = event.target.files[0];
        const delimiter = file.name.endsWith('.tsv') ? '\t' : ',';
        dataPromise = new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve({ data: e.target.result, delimiter: delimiter });
            reader.readAsText(file);
        });
    } else if (inputMode === 'raw') {
        const rawText = document.getElementById('rawData').value;
        if (!rawText.trim()) return;

        const commaCount = (rawText.match(/,/g) || []).length;
        const tabCount = (rawText.match(/\t/g) || []).length;
        const delimiter = tabCount > commaCount ? '\t' : ',';

        dataPromise = Promise.resolve({ data: rawText, delimiter: delimiter });
    } else {
        return;
    }

    dataPromise.then(result => {
        currentCsvData = result.data;
        currentDelimiter = result.delimiter;
        setFontControlsLock(true);
        redrawTable();
        exportButtons.style.display = 'flex';
    }).catch(error => {
        console.error("Error reading file or data:", error);
        tableContainer.innerHTML = '<p style="color: red;">Error processing data. Check formatting or encoding.</p>';
    });
}

function redrawTable() {
    if (!currentCsvData) return;

    const textColor = textColorInput.value;
    const borderColor = borderColorInput.value; // NEW: Get border color
    const tableHtml = generateApaTableHtml(currentCsvData, currentDelimiter, textColor, borderColor); // UPDATED: Pass border color
    tableContainer.innerHTML = tableHtml;
}


function generateApaTableHtml(csvText, delimiter, textColor, borderColor) { // UPDATED: Accept border color
    const fontFamily = fontFamilySelect.value;
    const fontSize = fontSizeSelect.value;
    const rawCaption = tableCaptionInput.value;
    const noteText = tableNoteTextarea.value;
    const rawAlignment = columnAlignmentInput.value;

    // Set a custom CSS variable on the table element
    const tableStyle = `font-family: '${fontFamily}'; font-size: ${fontSize}; color: ${textColor}; --apa-border-color: ${borderColor};`;
    const textStyle = `style="color: ${textColor}; font-family: '${fontFamily}'; font-size: ${fontSize};"`;

    const alignments = rawAlignment.toLowerCase()
        .split(',')
        .map(a => a.trim())
        .filter(a => ['left', 'center', 'right'].includes(a));

    let captionHtml = '';
    if (rawCaption) {
        const parts = rawCaption.split('.');
        const tableNumber = parts[0].trim();
        const tableTitle = parts.slice(1).join('.').trim();

        if (tableNumber) {
            captionHtml += `
                <p class="apa-caption" ${textStyle}>
                    <strong>${tableNumber}.</strong>
                    <span style="font-style: italic;">${tableTitle}</span>
                </p>`;
        }
    }

    let noteHtml = '';
    if (noteText) {
        noteHtml += `<p class="apa-note" ${textStyle}>Note. ${noteText}</p>`;
    }

    const rows = csvText.trim().split('\n').filter(row => row.trim() !== '');
    if (rows.length === 0) return '<p>No data found.</p>';
    const data = rows.map(row => row.split(delimiter).map(cell => cell.trim()));

    const numColumns = data[0].length;
    if (alignments.length !== numColumns) {
        alignments.length = 0;
        alignments.push('left');
        for (let i = 1; i < numColumns; i++) {
            alignments.push('center');
        }
    }

    let tableHtml = `<table class="apa-table" id="apaTable" style="${tableStyle}">`;

    const createCellHtml = (cellData, index, tag) => {
        const alignment = alignments[index];
        const alignmentStyle = `text-align: ${alignment};`;
        const wrapStyle = tag === 'th' ? `white-space: normal;` : '';
        return `<${tag} style="${alignmentStyle} ${wrapStyle}">${cellData}</${tag}>`;
    };

    const header = data[0];
    tableHtml += '<thead><tr>';
    header.forEach((cell, index) => {
        tableHtml += createCellHtml(cell, index, 'th');
    });
    tableHtml += '</tr></thead>';

    const body = data.slice(1);
    tableHtml += '<tbody>';
    body.forEach(row => {
        if (row.length === numColumns) {
            tableHtml += '<tr>';
            row.forEach((cell, index) => {
                tableHtml += createCellHtml(cell, index, 'td');
            });
            tableHtml += '</tr>';
        }
    });
    tableHtml += '</tbody>';
    tableHtml += '</table>';

    return captionHtml + tableHtml + noteHtml;
}

function exportToWord() {
    const contentToExport = tableContainer.innerHTML;
    if (!contentToExport.includes('apa-table')) {
        alert("Please generate a table first.");
        return;
    }
    const filename = 'APA_Table.doc';
    const borderColor = borderColorInput.value; // NEW: Get border color for export
    const textColor = textColorInput.value;
    const heavyBorderStyle = `2pt solid ${borderColor}`; // UPDATED: Use new variable
    const thinBorderStyle = `1pt solid ${borderColor}`; // UPDATED: Use new variable

    const content = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'></head>
            <body>
                <style>
                    /* Styles simplified for DOC export using user's color */
                    .apa-table { border-collapse: collapse; width: 100%; border: none; line-height: 1.5; color: ${textColor}; }
                    .apa-table th, .apa-table td { border: none; padding: 4px 10px; }
                    .apa-table thead { border-top: ${heavyBorderStyle}; border-bottom: ${thinBorderStyle}; } // UPDATED
                    .apa-table tbody tr:last-child { border-bottom: ${heavyBorderStyle}; } // UPDATED
                    .apa-caption { color: ${textColor}; }
                    .apa-caption strong { font-weight: bold; display: block; }
                    .apa-caption span { font-style: italic; display: block; }
                    .apa-note { color: ${textColor}; font-style: normal; }
                </style>
                ${contentToExport}
            </body>
        </html>
    `;

    const blob = new Blob([content], {
        type: 'application/msword;charset=utf-8;'
    });

    saveAs(blob, filename);
}

function exportToPng() {
    if (!tableContainer.innerHTML.includes('apa-table')) {
        alert("Please generate a table first.");
        return;
    }

    html2canvas(tableContainer, {
        scale: 2,
        backgroundColor: '#ffffff',

        windowWidth: tableContainer.scrollWidth,
        windowHeight: tableContainer.scrollHeight
    }).then(canvas => {
        const link = document.createElement('a');
        link.download = 'APA_Table.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });
}

function copyPngToClipboard() {
    if (!tableContainer.innerHTML.includes('apa-table')) {
        alert("Please generate a table first.");
        return;
    }

    html2canvas(tableContainer, {
        scale: 2,
        backgroundColor: '#ffffff',

        windowWidth: tableContainer.scrollWidth,
        windowHeight: tableContainer.scrollHeight
    }).then(canvas => {
        if (navigator.clipboard && navigator.clipboard.write) {
            canvas.toBlob(blob => {
                const item = new ClipboardItem({ 'image/png': blob });
                navigator.clipboard.write([item]).then(() => {
                    alert('Formatted table image copied to clipboard!');
                }, (error) => {
                    console.error('Failed to copy to clipboard:', error);
                    alert('Failed to copy image. Please try the Download PNG option.');
                });
            }, 'image/png');
        } else {
            alert('Clipboard API not fully supported. Please use the Download PNG option.');
        }
    });
}