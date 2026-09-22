const fileInput = document.getElementById('fileInput');
const tableContainer = document.getElementById('tableContainer');
const exportBtn = document.getElementById('exportBtn');
const exportTechBtn = document.getElementById('exportTechBtn');
const clearBtn = document.getElementById('clearBtn');

let rawData = [];

// Event Listeners
fileInput.addEventListener('change', (e) => processSelectedFile(e.target.files[0]), false);
exportBtn.addEventListener('click', () => handleExport(false), false);
exportTechBtn.addEventListener('click', () => handleExport(true), false);
clearBtn.addEventListener('click', clearPreview, false);

// Drag & Drop Handling
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    tableContainer.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    tableContainer.addEventListener(eventName, () => tableContainer.classList.add('drag-over'), false);
});

['dragleave', 'drop'].forEach(eventName => {
    tableContainer.addEventListener(eventName, () => tableContainer.classList.remove('drag-over'), false);
});

tableContainer.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const file = dt.files[0];
    if (file) processSelectedFile(file);
});

function isHyperlinkObj(v) { 
    return v && typeof v === "object" && !!v.hyperlink; 
}

function parseHyperlinkFormula(f) {
    if (!f) return null;
    const m = /HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)/i.exec(f);
    if (m) return { hyperlink: m[1], text: m[2] || m[1] };
    return null;
}

function getFormattedTodayDate() {
    const today = new Date();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const year = String(today.getFullYear()).slice(-2);
    return `${month}/${day}/${year}`;
}

// Sanitization ONLY for Linked Request ID and Title fields
function sanitizeLinkedValue(val) {
    if (!val) return "";
    const str = String(val).trim();
    const lower = str.toLowerCase();
    
    if (lower === "not assigned" || lower === "https://techassist.bounty.org.ph/app/itdesk/ui/requests/-1/details") {
        return "";
    }
    return str;
}

// Check if any row has valid data specifically for a given column
function hasValidDataForColumn(data, colKey) {
    return data.some(row => {
        const val = row[colKey];
        if (!val) return false;
        if (isHyperlinkObj(val)) {
            return !!(val.text || val.hyperlink);
        }
        return !!val;
    });
}

function processSelectedFile(f) {
    if (!f) return;
    const reader = new FileReader();

    reader.onload = function(ev) {
        try {
            const data = new Uint8Array(ev.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            if (!sheet || !sheet['!ref']) return;

            const range = XLSX.utils.decode_range(sheet['!ref']);
            let headerRowIndex = -1;

            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    const val = cell ? String(XLSX.utils.format_cell(cell)).trim().toLowerCase() : "";
                    if (val === "department") {
                        headerRowIndex = r;
                        break;
                    }
                }
                if (headerRowIndex !== -1) break;
            }

            if (headerRowIndex === -1) headerRowIndex = range.s.r;

            const headersRaw = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
                headersRaw.push(cell ? XLSX.utils.format_cell(cell).trim() : `Column${c + 1}`);
            }

            const jsonData = [];
            let currentDepartment = "";

            for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
                const rowObj = {};
                let rowHasValue = false;

                for (let c = range.s.c; c <= range.e.c; c++) {
                    let colName = headersRaw[c - range.s.c];

                    // Standard Name Replacements
                    if (colName === "Change ID Caused By Request") colName = "Change ID";
                    if (colName === "Change Title Caused By Request" || colName === "Subject") colName = "Description";
                    
                    // Sprint Number Mapping
                    if (colName === "Sprint" || colName === "Sprint #" || colName === "Sprint No" || colName === "Sprint No.") {
                        colName = "Sprint Number";
                    }

                    // Special Release Request ID & Title Replacements
                    if (colName === "Linked Request ID") colName = "Special Release Request ID";
                    if (colName === "Linked Request Title") colName = "Special Release Request Title";

                    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                    if (!cell) { rowObj[colName] = ""; continue; }

                    let text = XLSX.utils.format_cell(cell);
                    let linkObj = (cell.l && cell.l.Target) 
                        ? { text: text || cell.l.Target, hyperlink: cell.l.Target } 
                        : parseHyperlinkFormula(cell.f);

                    // Apply "Not Assigned" / -1 URL filter ONLY to Special Release fields
                    const isSpecialCol = (colName === "Special Release Request ID" || colName === "Special Release Request Title");

                    if (linkObj) {
                        const cleanText = isSpecialCol ? sanitizeLinkedValue(linkObj.text) : (linkObj.text ? String(linkObj.text).trim() : "");
                        const cleanLink = isSpecialCol ? sanitizeLinkedValue(linkObj.hyperlink) : (linkObj.hyperlink ? String(linkObj.hyperlink).trim() : "");

                        if (!cleanText && !cleanLink) {
                            rowObj[colName] = "";
                        } else {
                            rowObj[colName] = { text: cleanText || cleanLink, hyperlink: cleanLink };
                            rowHasValue = true;
                        }
                    } else {
                        const cleanVal = isSpecialCol ? sanitizeLinkedValue(text) : (text ? String(text).trim() : "");
                        rowObj[colName] = cleanVal;
                        if (cleanVal) rowHasValue = true;
                    }
                }

                // Ensure Sprint Number property exists even if missing in original uploaded file
                if (!rowObj.hasOwnProperty('Sprint Number')) {
                    rowObj['Sprint Number'] = "";
                }

                // Fallback Logic: If RequestID or Description are empty, fallback to Special Release Request values
                if (!rowObj['RequestID'] && rowObj['Special Release Request ID']) {
                    rowObj['RequestID'] = rowObj['Special Release Request ID'];
                }
                if (!rowObj['Description'] && rowObj['Special Release Request Title']) {
                    rowObj['Description'] = rowObj['Special Release Request Title'];
                }

                const deptVal = rowObj["Department"];
                const parsedDept = isHyperlinkObj(deptVal) ? deptVal.text : String(deptVal || "").trim();
                if (parsedDept) {
                    currentDepartment = parsedDept;
                } else {
                    rowObj["Department"] = currentDepartment;
                }

                if (rowHasValue) {
                    jsonData.push(rowObj);
                }
            }

            rawData = jsonData;
            renderEditableTable(rawData);

            exportBtn.disabled = false;
            exportTechBtn.disabled = false;
            clearBtn.disabled = false;
        } catch (err) {
            console.error(err);
            alert('Failed to parse file.');
        }
    };
    reader.readAsArrayBuffer(f);
}

function getActiveHeaders(data, includeDepartment = true, includeTechnician = true) {
    const hasSpecialTitle = hasValidDataForColumn(data, 'Special Release Request Title');
    const hasSpecialID = hasValidDataForColumn(data, 'Special Release Request ID');
    const includeSpecial = hasSpecialTitle || hasSpecialID;

    const baseHeaders = [];

    if (includeDepartment) baseHeaders.push('Department');
    
    // Always include 'Sprint Number' so users can fill it manually
    baseHeaders.push('Sprint Number', 'Change Type', 'RequestID', 'Change ID', 'Description');

    if (includeSpecial) {
        baseHeaders.push('Special Release Request ID', 'Special Release Request Title');
    }

    if (includeTechnician) baseHeaders.push('Technician');

    baseHeaders.push('Requester', 'UAT Owner', 'UAT Date', 'Request Status');

    return baseHeaders;
}

function renderEditableTable(data) {
    tableContainer.innerHTML = "";

    const headers = getActiveHeaders(data, true, true);

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerTr = document.createElement('tr');
    
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        headerTr.appendChild(th);
    });
    thead.appendChild(headerTr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    data.forEach((row) => {
        const tr = document.createElement('tr');

        headers.forEach(header => {
            const td = document.createElement('td');
            td.contentEditable = "true";
            
            const cellVal = row[header];

            if (isHyperlinkObj(cellVal)) {
                td.innerHTML = `<a href="${cellVal.hyperlink}" target="_blank">${cellVal.text || cellVal.hyperlink}</a>`;
            } else {
                td.textContent = cellVal || "";
            }

            td.addEventListener('blur', () => {
                const updatedText = td.textContent.trim();
                if (isHyperlinkObj(row[header])) {
                    row[header].text = updatedText;
                } else {
                    row[header] = updatedText;
                }
            });

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableContainer.appendChild(table);
}

function groupByDepartment(rows) {
    const groups = [];
    let currentDept = null;
    let currentRows = [];

    for (const r of rows) {
        const deptVal = r["Department"];
        const dept = isHyperlinkObj(deptVal) ? (deptVal.text || "") : String(deptVal || "");

        if (!dept.trim()) {
            if (currentDept) currentRows.push(r);
            continue;
        }

        if (!currentDept || dept !== currentDept) {
            if (currentDept) groups.push({ department: currentDept, rows: currentRows });
            currentDept = dept;
            currentRows = [r];
        } else {
            currentRows.push(r);
        }
    }
    if (currentDept) groups.push({ department: currentDept, rows: currentRows });
    return groups;
}

async function handleExport(includeTechnician) {
    if (!rawData.length) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Announcement');

    const purpleHeaderFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF5451E0' }
    };

    const whiteBoldFont = {
        name: 'Segoe UI',
        bold: true,
        color: { argb: 'FFFFFFFF' },
        size: 10
    };

    const lightBorder = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    const exportDateStr = getFormattedTodayDate();
    const customWidths = {
        'Department': 25, 'Sprint Number': 16, 'Change Type': 15, 'RequestID': 16,
        'Change ID': 14, 'Description': 65, 'Special Release Request ID': 22,
        'Special Release Request Title': 42, 'Technician': 22,
        'Requester': 22, 'UAT Owner': 22, 'UAT Date': 14, 'Request Status': 16
    };

    if (includeTechnician) {
        const exportHeaders = getActiveHeaders(rawData, true, true);

        const hRow = sheet.addRow(exportHeaders);
        hRow.height = 24;
        hRow.eachCell(c => {
            c.font = whiteBoldFont;
            c.fill = purpleHeaderFill;
            c.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        rawData.forEach(row => {
            const exportRow = { ...row };
            const reqVal = exportRow['Requester'];
            exportRow['UAT Owner'] = isHyperlinkObj(reqVal) ? (reqVal.text || reqVal.hyperlink) : reqVal;
            exportRow['UAT Date'] = exportDateStr;

            const rowValues = exportHeaders.map(h => {
                const val = exportRow[h];
                if (isHyperlinkObj(val)) return val.text || val.hyperlink;
                return val || "";
            });

            const r = sheet.addRow(rowValues);
            r.height = 20;

            r.eachCell({ includeEmpty: true }, (cell, colIdx) => {
                const colKey = exportHeaders[colIdx - 1];
                const rawVal = exportRow[colKey];

                cell.font = { name: 'Segoe UI', size: 10 };
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
                cell.border = lightBorder;

                if (isHyperlinkObj(rawVal)) {
                    cell.value = { text: rawVal.text || rawVal.hyperlink, hyperlink: rawVal.hyperlink };
                    cell.font = { name: 'Segoe UI', size: 10, color: { argb: "FF0000FF" }, underline: true };
                }
            });
        });

        sheet.columns.forEach((col, idx) => {
            const headerName = exportHeaders[idx];
            col.width = customWidths[headerName] || 20;
        });

    } else {
        const exportHeaders = getActiveHeaders(rawData, false, false);

        const hRow = sheet.addRow(exportHeaders);
        hRow.height = 24;
        hRow.eachCell(c => {
            c.font = whiteBoldFont;
            c.fill = purpleHeaderFill;
            c.alignment = { vertical: 'middle', horizontal: 'left' };
        });

        const grouped = groupByDepartment(rawData);

        grouped.forEach(group => {
            const dRow = sheet.addRow([group.department]);
            dRow.height = 20;
            sheet.mergeCells(dRow.number, 1, dRow.number, exportHeaders.length);
            
            const deptCell = dRow.getCell(1);
            deptCell.font = whiteBoldFont;
            deptCell.fill = purpleHeaderFill;
            deptCell.alignment = { vertical: 'middle', horizontal: 'left' };

            group.rows.forEach(row => {
                const exportRow = { ...row };
                const reqVal = exportRow['Requester'];
                exportRow['UAT Owner'] = isHyperlinkObj(reqVal) ? (reqVal.text || reqVal.hyperlink) : reqVal;
                exportRow['UAT Date'] = exportDateStr;

                const rowValues = exportHeaders.map(h => {
                    const val = exportRow[h];
                    if (isHyperlinkObj(val)) return val.text || val.hyperlink;
                    return val || "";
                });

                const r = sheet.addRow(rowValues);
                r.height = 20;

                r.eachCell({ includeEmpty: true }, (cell, colIdx) => {
                    const colKey = exportHeaders[colIdx - 1];
                    const rawVal = exportRow[colKey];

                    cell.font = { name: 'Segoe UI', size: 10 };
                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.border = lightBorder;

                    if (isHyperlinkObj(rawVal)) {
                        cell.value = { text: rawVal.text || rawVal.hyperlink, hyperlink: rawVal.hyperlink };
                        cell.font = { name: 'Segoe UI', size: 10, color: { argb: "FF0000FF" }, underline: true };
                    }
                });
            });
        });

        sheet.columns.forEach((col, idx) => {
            const headerName = exportHeaders[idx];
            col.width = customWidths[headerName] || 20;
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = includeTechnician ? 'Release_Template_with_Tech.xlsx' : 'Release_Template_detailed.xlsx';
    saveAs(new Blob([buffer]), fileName);
}

function clearPreview() {
    rawData = []; 
    fileInput.value = "";
    tableContainer.innerHTML = `
      <div class="placeholder" id="dropZone">
        <div class="placeholder-icon">📂</div>
        <p>Upload or drag & drop a file here to preview its contents.</p>
      </div>`;
    exportBtn.disabled = true;
    exportTechBtn.disabled = true;
    clearBtn.disabled = true;
}