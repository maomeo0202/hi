/******************** CONFIG *************************/
const SHEET_ID = '1Pmj5slDmLxzNBxY3_n4IVVm48cSqOlrctug2YxWs2Ls';
const FOLDER_ID = '1aP70_1o6X92PqfCNry72CBav3ExNnjuJ';
const IMG_FOLDER_ID = '1xI8URRci3VXc3keF6CNliwwoaXJgKVMb'; // Folder for images
const ADMIN_SHEET_NAME = 'ADMIN_USERS';
const DB_SHEET_NAME = 'DATABASE_NOTE';
const PAGE_SIZE = 10;


/******************** LOAD APP *************************/
function doGet() {
    return HtmlService.createTemplateFromFile('trang_chu')
        .evaluate()
        .setTitle('Note App SPA')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// /******************** SYS ACCOUNT CHECK *************************/
// function isSysAccount(username, password) {
//   return String(username).trim() === "sys" && String(password).trim() === "a1";
// }
// hàm đăng nhập offline nhưng chưa được phân quyền 
/******************** LOGIN *************************/
function login(username, password, page) {
    // if (isSysAccount(username,password)){
    //   return{
    //     success: true,
    //     info: {
    //       user: "sys",
    //       canAdd: true,
    //       canEdit: true,
    //       canDelete: true,
    //       access: {
    //         trang_chu: true,
    //         app_taive: true,
    //         bo_sung: true,
    //         quan_tri: true
    //       }
    //     }
    //   };
    // }


    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    let data = sh.getDataRange().getValues();
    data.shift();

    for (let r of data) {
        if (String(r[0]).trim() === String(username).trim() &&
            String(r[1]).trim() === String(password).trim()) {

            return {
                success: true,
                info: {
                    user: r[0],
                    canAdd: r[3],
                    canEdit: r[4],
                    canDelete: r[5],
                    access: {
                        trang_chu: r[6],
                        app_taive: r[7],
                        bo_sung: r[8],
                        quan_tri: r[9]
                    }
                }
            };
        }
    }
    return {
        success: false,
        message: "Sai tài khoản hoặc mật khẩu."
    };
}

/******************** CHECK ACCESS *************************/
function checkAccess(username, page) {
    // if (String(username).trim() === "sys") return true;
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    let data = sh.getDataRange().getValues();
    data.shift();

    let map = {
        trang_chu: 6,
        app_taive: 7,
        bo_sung: 8,
        quan_tri: 9
    };
    let col = map[page];

    for (let r of data) {
        if (String(r[0]) === String(username)) return !!r[col];
    }
    return false;
}

/******************** UPLOAD FILE *************************/
function uploadFile(fileObj) {
    let base64 = fileObj.base64.split("base64,")[1];
    let bytes = Utilities.base64Decode(base64);
    let blob = Utilities.newBlob(bytes, fileObj.mimeType, fileObj.fileName);
    let folder = DriveApp.getFolderById(FOLDER_ID);
    let file = folder.createFile(blob);
    return {
        url: file.getUrl(),
        fileName: file.getName()
    };
}

/******************** UPLOAD IMAGE *************************/
function uploadImage(fileObj) {
    let base64 = fileObj.base64.split("base64,")[1];
    let bytes = Utilities.base64Decode(base64);
    let blob = Utilities.newBlob(bytes, fileObj.mimeType, fileObj.fileName);
    let folder = DriveApp.getFolderById(IMG_FOLDER_ID);
    let file = folder.createFile(blob);
    return {
        url: file.getUrl()
    };
}

/******************** UPLOAD IMAGE FROM BASE64 *************************/
function uploadImageFromBase64(base64Data, mimeType) {
    try {
        let bytes = Utilities.base64Decode(base64Data);
        let blob = Utilities.newBlob(bytes, mimeType, 'pasted-image.' + (mimeType.split('/')[1] || 'png'));
        let folder = DriveApp.getFolderById(IMG_FOLDER_ID);
        let file = folder.createFile(blob);
        return {
            url: file.getUrl()
        };
    } catch (e) {
        return {
            error: e.toString()
        };
    }
}

/******************** UPDATE NOTE COLUMN *************************/
function updateNoteColumn(rowNum, colName, value) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                success: false,
                error: "Sheet not found"
            };
        }

        const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
        const colIndex = headers.indexOf(colName);

        if (colIndex === -1) {
            return {
                success: false,
                error: "Column not found: " + colName
            };
        }

        sh.getRange(rowNum, colIndex + 1).setValue(value);

        return {
            success: true
        };
    } catch (e) {
        return {
            success: false,
            error: e.toString()
        };
    }
}

/******************** GENERATE UNIQUE ID *************************/
function generateUniqueId() {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);

    // Get all existing IDs from the sheet
    let existingIds = [];
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf('ID');

    if (idColIndex !== -1 && sh.getLastRow() > 1) {
        // Get all IDs from column (skip header row)
        const idRange = sh.getRange(2, idColIndex + 1, sh.getLastRow() - 1, 1);
        const idValues = idRange.getValues();
        existingIds = idValues.map(row => String(row[0]).trim()).filter(id => id !== '');
    }

    // Generate a random 8-digit ID
    let newId;
    let attempts = 0;
    const maxAttempts = 100;

    do {
        // Generate random 8-digit number (10000000 to 99999999)
        newId = String(Math.floor(10000000 + Math.random() * 90000000));
        attempts++;

        // Safety check to prevent infinite loop
        if (attempts >= maxAttempts) {
            // If we can't find unique ID after many attempts, use timestamp-based ID
            const timestamp = String(Date.now()).slice(-8);
            newId = timestamp.padStart(8, '0');
            break;
        }
    } while (existingIds.includes(newId));

    return newId;
}

/******************** SAVE NOTE *************************/
function saveNote(note, username) {
    try {
        let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);

        if (!sh) {
            return {
                success: false,
                error: 'Sheet not found: ' + DB_SHEET_NAME
            };
        }

        // Check if sheet is empty (no header row)
        if (sh.getLastRow() === 0) {
            sh.appendRow([
                "ID", "Time", "Lĩnh vực", "Địa bàn",
                "Tiêu đề", "Chi tiết", "Hướng GQ", "Góp ý",
                "img", "File", "User", "log_update"
            ]);
        } else {
            // Sheet exists, check if ID column exists
            const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
            const idColIndex = headers.indexOf('ID');

            if (idColIndex === -1) {
                // ID column doesn't exist, add it as the first column
                sh.insertColumnBefore(1);
                sh.getRange(1, 1).setValue('ID');

                // Fill existing rows with generated IDs
                const lastRow = sh.getLastRow();
                if (lastRow > 1) {
                    for (let row = 2; row <= lastRow; row++) {
                        const existingId = sh.getRange(row, 1).getValue();
                        if (!existingId || String(existingId).trim() === '') {
                            sh.getRange(row, 1).setValue(generateUniqueId());
                        }
                    }
                }
            }
        }

        // Generate unique 8-digit ID
        const uniqueId = generateUniqueId();

        // Get current headers to ensure correct column order
        const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

        // Debug: log headers
        Logger.log('Headers: ' + JSON.stringify(headers));
        Logger.log('Note data: ' + JSON.stringify(note));

        // Get imgUrls (multiple images separated by comma and newline)
        const imgUrls = note.imgUrls || '';

        // Create date in Vietnam timezone (GMT+7)
        // Get current UTC time and add 7 hours for Vietnam timezone
        const now = new Date();
        const utcTime = now.getTime() + (now.getTimezoneOffset() * 60 * 1000);
        const vietnamTime = new Date(utcTime + (7 * 60 * 60 * 1000)); // Add 7 hours for GMT+7

        // Build row array based on header order
        const newRow = [];
        headers.forEach(header => {
            const headerStr = String(header).trim();
            const headerLower = headerStr.toLowerCase();

            if (headerStr === 'ID') {
                newRow.push(uniqueId);
            } else if (headerStr === 'Time' || headerLower === 'time') {
                newRow.push(vietnamTime);
            } else if (headerLower.includes('lĩnh vực') || headerLower.includes('linh vuc') || headerLower === 'linh_vuc') {
                newRow.push(note.linhVuc || '');
            } else if (headerLower.includes('địa bàn') || headerLower.includes('dia ban') || headerLower === 'dia_ban') {
                newRow.push(note.diaBan || '');
            } else if (headerLower.includes('tiêu đề') || headerLower.includes('tieu de') || headerLower === 'tieu_de') {
                newRow.push(note.tieuDe || '');
            } else if (headerLower.includes('chi tiết') || headerLower.includes('chi tiet') || headerLower === 'chi_tiet') {
                newRow.push(note.chiTiet || '');
            } else if ((headerLower.includes('hướng') && headerLower.includes('gq')) || headerLower.includes('huong gq') || headerLower === 'huong_gq') {
                newRow.push(note.huongGQ || '');
            } else if (headerLower.includes('góp ý') || headerLower.includes('gop y') || headerLower === 'gop_y') {
                newRow.push(note.gopY || '');
            } else if (headerLower === 'img' || (headerLower.includes('img') && !headerLower.includes('file'))) {
                newRow.push(imgUrls);
            } else if ((headerLower === 'file' || headerLower.includes('file')) && !headerLower.includes('img')) {
                newRow.push(note.fileUrl || '');
            } else if (headerLower === 'user') {
                newRow.push(username);
            } else if (headerLower === 'log_update' || headerLower.includes('log_update')) {
                newRow.push(''); // log_update - empty for new notes
            } else {
                newRow.push(''); // Default empty for unknown columns
            }
        });

        // Validate that newRow has the same length as headers
        if (newRow.length !== headers.length) {
            Logger.log('Row length mismatch: expected ' + headers.length + ' columns, got ' + newRow.length);
            return {
                success: false,
                error: 'Row length mismatch: expected ' + headers.length + ' columns, got ' + newRow.length
            };
        }

        try {
            sh.appendRow(newRow);
            Logger.log('Row appended successfully. New row: ' + JSON.stringify(newRow));
            return {
                success: true
            };
        } catch (e) {
            Logger.log('Error appending row: ' + e.toString());
            return {
                success: false,
                error: e.toString()
            };
        }
    } catch (e) {
        Logger.log('Error in saveNote: ' + e.toString());
        return {
            success: false,
            error: e.toString()
        };
    }
}

/******************** GET NOTES *************************/
// function getNotes(page) {
//   let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
//   let v = sh.getDataRange().getValues();

//   if (v.length < 2) return { headers:[], rows:[] };

//   let headers = v.shift();
//   v.reverse();

//   let total = v.length;
//   let pages = Math.ceil(total / PAGE_SIZE);
//   page = Math.max(1, Math.min(page, pages));

//   let rows = v.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

//   return { headers, rows, page, totalPages: pages };
// }
// function getNotes(page) {
//   let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
//   let v = sh.getDataRange().getValues();

//   if (v.length < 2) return { headers:[], rows:[], page:1, totalPages:1 };

//   let headers = v.shift();     // lấy hàng 1
//   let data = v.reverse();      // đảo dữ liệu (không đảo header)

//   let total = data.length;
//   let pages = Math.ceil(total / PAGE_SIZE);

//   page = Math.max(1, Math.min(page, pages));

//   let rows = data.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

//   // Chuyển từng dòng thành object
//   rows = rows.map(r => {
//     let o = {};
//     headers.forEach((h, i) => o[h] = r[i]);
//     return o;
//   });

//   return { headers, rows, page, totalPages: pages };
// }

// function getNotes(page) {
//   console.log("=== DEBUG getNotes(page) ===");
//   console.log("Page nhận vào:", page);

//   try {

//     console.log("DATABASE_NOTES hiện tại:", DATABASE_NOTES);

//     if (!DATABASE_NOTES || !Array.isArray(DATABASE_NOTES)) {
//       console.log("DATABASE_NOTES bị undefined hoặc không phải mảng!");
//       return { list: [], totalPages: 0 };
//     }

//     // debug từng row
//     DATABASE_NOTES.forEach((r, i) => console.log("Row", i, r));

//     // nếu có phân trang
//     const perPage = 20;
//     const start = (page - 1) * perPage;
//     const end = start + perPage;

//     const list = DATABASE_NOTES.slice(start, end);
//     console.log("List sau phân trang:", list);

//     return {
//       list: list,
//       totalPages: Math.ceil(DATABASE_NOTES.length / perPage)
//     };

//   } catch (e) {
//     console.log("Lỗi getNotes:", e);
//     return { list: [], totalPages: 0 };
//   }
// }

function getNotes(page, pageSizeParam) {
    try {

        // đảm bảo page luôn có giá trị số >=1
        if (typeof page === 'undefined' || page === null || isNaN(Number(page))) {
            page = 1;
        } else {
            page = Number(page);
        }

        // Get page size (default to PAGE_SIZE if not provided)
        let pageSize = PAGE_SIZE;
        if (typeof pageSizeParam !== 'undefined' && pageSizeParam !== null && !isNaN(Number(pageSizeParam))) {
            pageSize = Number(pageSizeParam);
        }


        // Kiểm tra SHEET_ID và DB_SHEET_NAME có tồn tại không
        if (!SHEET_ID || !DB_SHEET_NAME) {
            const errMsg = "SHEET_ID hoặc DB_SHEET_NAME chưa được cấu hình";
            return {
                headers: [],
                rows: [],
                page: 1,
                totalPages: 0,
                error: errMsg
            };
        }

        let ss;
        try {
            ss = SpreadsheetApp.openById(SHEET_ID);
        } catch (e) {
            return {
                headers: [],
                rows: [],
                page: page,
                totalPages: 0,
                error: "Không thể truy cập Spreadsheet. Kiểm tra SHEET_ID và quyền truy cập."
            };
        }

        if (!ss) {
            return {
                headers: [],
                rows: [],
                page: page,
                totalPages: 0,
                error: "Spreadsheet not found"
            };
        }

        const sh = ss.getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                headers: [],
                rows: [],
                page: page,
                totalPages: 0,
                error: "Sheet '" + DB_SHEET_NAME + "' not found"
            };
        }

        const v = sh.getDataRange().getValues();

        if (v.length < 2) {
            return {
                headers: v[0] || [],
                rows: [],
                page: page,
                totalPages: 0
            };
        }

        const headers = v.shift(); // header row
        const totalRows = v.length;

        // Store original row numbers before reversing
        const dataWithRowNum = v.map((row, idx) => ({
            rowNum: idx + 2, // +2 because row 1 is header, data starts at row 2
            data: row
        }));

        // Reverse to show newest first
        dataWithRowNum.reverse();

        const total = dataWithRowNum.length;
        const pages = Math.max(1, Math.ceil(total / pageSize));
        page = Math.max(1, Math.min(page, pages));

        const rowsSlice = dataWithRowNum.slice((page - 1) * pageSize, page * pageSize);

        // Helper function to convert values to serializable format
        function serializeValue(val) {
            if (val === null || val === undefined) {
                return '';
            }
            // Convert Date objects to Vietnam timezone string
            if (val instanceof Date) {
                return Utilities.formatDate(val, "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss");
            }
            // Convert arrays
            if (Array.isArray(val)) {
                return val.map(serializeValue);
            }
            // Return primitive values as-is
            return val;
        }

        // convert each row array -> object keyed by headers, include rowNum
        const rows = rowsSlice.map(item => {
            const o = {
                _rowNum: item.rowNum
            }; // Store original spreadsheet row number
            headers.forEach((h, i) => {
                o[h] = serializeValue(item.data[i]);
            });
            return o;
        });

        // Ensure headers are also serializable (convert any Date objects)
        const serializedHeaders = headers.map(h => serializeValue(h));

        const result = {
            headers: serializedHeaders,
            rows: rows,
            page: page,
            totalPages: pages,
            totalRows: total
        };

        // Test serialization before returning
        try {
            JSON.stringify(result);
        } catch (serialError) {
            return {
                headers: [],
                rows: [],
                page: 1,
                totalPages: 0,
                error: "Serialization error: " + serialError.toString()
            };
        }

        return result;
    } catch (e) {
        const errorResult = {
            headers: [],
            rows: [],
            page: 1,
            totalPages: 0,
            error: e.toString()
        };
        return errorResult;
    }
}



// function getNotes(limit) {
//   Logger.log("DEBUG: getNotes() chạy");

//   const ss = SpreadsheetApp.openById(SHEET_ID);
//   const sh = ss.getSheetByName(DB_SHEET_NAME);

//   if (!sh) {
//     Logger.log("ERROR: Sheet NOTE không tồn tại");
//     return { headers: [], rows: [] };
//   }

//   const data = sh.getDataRange().getValues();

//   if (data.length < 2) {
//     Logger.log("DEBUG: không có dữ liệu (chỉ header)");
//     return { headers: data[0], rows: [] };
//   }

//   const headers = data[0];
//   const rows = [];

//   for (let i = 1; i < data.length; i++) {
//     let obj = {};
//     headers.forEach((h, idx) => {
//       obj[h] = data[i][idx];
//     });
//     rows.push(obj);
//   }

//   Logger.log("DEBUG: trả về rows = " + JSON.stringify(rows));
//   return { headers, rows };
// }






/******************** GET NOTE BY ROW *************************/
function getNoteByRow(rowNum) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return null;
        }

        const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
        const rowData = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];

        const note = {};
        headers.forEach((h, i) => {
            note[h] = rowData[i];
        });

        // Map to expected field names
        return {
            linhVuc: note['Lĩnh vực'] || note['Linh vuc'] || '',
            diaBan: note['Địa bàn'] || note['Dia ban'] || '',
            tieuDe: note['Tiêu đề'] || note['Tieu de'] || note['Tiêu đề / Vấn đề'] || '',
            chiTiet: note['Chi tiết'] || note['Chi tiet'] || note['Chi tiết vấn đề'] || '',
            huongGQ: note['Hướng GQ'] || note['Huong GQ'] || note['Hướng giải quyết'] || '',
            gopY: note['Góp ý'] || note['Gop y'] || note['Góp ý hướng giải quyết'] || '',
            fileUrl: note['File'] || '',
            imgUrls: note['img'] || note['Img'] || ''
        };
    } catch (e) {
        return null;
    }
}

// Helper function to format Vietnam timezone date
function formatVietnamDateTime(date) {
    // Get current UTC time and add 7 hours for Vietnam timezone
    const utcTime = date.getTime() + (date.getTimezoneOffset() * 60 * 1000);
    const vietnamTime = new Date(utcTime + (7 * 60 * 60 * 1000)); // Add 7 hours for GMT+7

    const year = vietnamTime.getFullYear();
    const month = String(vietnamTime.getMonth() + 1).padStart(2, '0');
    const day = String(vietnamTime.getDate()).padStart(2, '0');
    const hours = String(vietnamTime.getHours()).padStart(2, '0');
    const minutes = String(vietnamTime.getMinutes()).padStart(2, '0');
    const seconds = String(vietnamTime.getSeconds()).padStart(2, '0');

    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/******************** UPDATE NOTE *************************/
function updateNote(rowNum, note, username) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                success: false,
                error: "Sheet not found"
            };
        }

        // Get headers to find column indices
        let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

        // Debug: log headers and note data
        Logger.log('updateNote - Headers: ' + JSON.stringify(headers));
        Logger.log('updateNote - Note data: ' + JSON.stringify(note));
        Logger.log('updateNote - RowNum: ' + rowNum);

        // Check if log_update column exists, if not add it
        let logUpdateColIndex = headers.indexOf('log_update');
        if (logUpdateColIndex === -1) {
            // Add log_update column at the end
            const lastCol = sh.getLastColumn();
            sh.getRange(1, lastCol + 1).setValue('log_update');
            headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
            logUpdateColIndex = headers.indexOf('log_update');
        }

        // Update row data
        const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];

        // Helper function to find column index by matching header variations
        const findColIndex = (possibleNames) => {
            for (let name of possibleNames) {
                const exactIndex = headers.indexOf(name);
                if (exactIndex !== -1) {
                    return exactIndex;
                }
            }
            // If exact match not found, try case-insensitive and partial match
            for (let i = 0; i < headers.length; i++) {
                const h = String(headers[i]).trim();
                const hLower = h.toLowerCase();
                for (let name of possibleNames) {
                    const nameLower = String(name).toLowerCase();
                    if (hLower === nameLower ||
                        hLower.includes(nameLower) ||
                        nameLower.includes(hLower)) {
                        return i;
                    }
                }
            }
            return -1;
        };

        // Find and update each field
        const linhVucCols = ['Lĩnh vực', 'Linh vuc', 'linh_vuc', 'Linh_vuc'];
        const diaBanCols = ['Địa bàn', 'Dia ban', 'dia_ban', 'Dia_ban'];
        const tieuDeCols = ['Tiêu đề', 'Tieu de', 'tieu_de', 'Tieu_de', 'Tiêu đề / Vấn đề'];
        const chiTietCols = ['Chi tiết', 'Chi tiet', 'chi_tiet', 'Chi_tiet', 'Chi tiết vấn đề'];
        const huongGQCols = ['Hướng GQ', 'Huong GQ', 'huong_gq', 'Huong_gq', 'Hướng giải quyết'];
        const gopYCols = ['Góp ý', 'Gop y', 'gop_y', 'Gop_y', 'Góp ý hướng giải quyết'];
        const fileCols = ['File', 'file'];
        const imgCols = ['img', 'Img'];

        // Update each field if column found
        const linhVucIdx = findColIndex(linhVucCols);
        if (linhVucIdx !== -1) {
            row[linhVucIdx] = note.linhVuc || '';
            Logger.log('Updated Lĩnh vực at index ' + linhVucIdx + ': ' + note.linhVuc);
        }

        const diaBanIdx = findColIndex(diaBanCols);
        if (diaBanIdx !== -1) {
            row[diaBanIdx] = note.diaBan || '';
            Logger.log('Updated Địa bàn at index ' + diaBanIdx + ': ' + note.diaBan);
        }

        const tieuDeIdx = findColIndex(tieuDeCols);
        if (tieuDeIdx !== -1) {
            row[tieuDeIdx] = note.tieuDe || '';
            Logger.log('Updated Tiêu đề at index ' + tieuDeIdx + ': ' + note.tieuDe);
        }

        const chiTietIdx = findColIndex(chiTietCols);
        if (chiTietIdx !== -1) {
            row[chiTietIdx] = note.chiTiet || '';
            Logger.log('Updated Chi tiết at index ' + chiTietIdx + ': ' + note.chiTiet);
        }

        const huongGQIdx = findColIndex(huongGQCols);
        if (huongGQIdx !== -1) {
            row[huongGQIdx] = note.huongGQ || '';
            Logger.log('Updated Hướng GQ at index ' + huongGQIdx + ': ' + note.huongGQ);
        }

        const gopYIdx = findColIndex(gopYCols);
        if (gopYIdx !== -1) {
            row[gopYIdx] = note.gopY || '';
            Logger.log('Updated Góp ý at index ' + gopYIdx + ': ' + note.gopY);
        }

        const fileIdx = findColIndex(fileCols);
        if (fileIdx !== -1) {
            row[fileIdx] = note.fileUrl || '';
            Logger.log('Updated File at index ' + fileIdx + ': ' + note.fileUrl);
        }

        const imgIdx = findColIndex(imgCols);
        if (imgIdx !== -1) {
            row[imgIdx] = note.imgUrls || '';
            Logger.log('Updated img at index ' + imgIdx + ': ' + note.imgUrls);
        }

        // Update log_update column
        const now = new Date();
        const logEntry = formatVietnamDateTime(now) + ' - ' + username;

        // Get existing log_update value
        let existingLog = '';
        if (logUpdateColIndex >= 0 && row[logUpdateColIndex]) {
            existingLog = String(row[logUpdateColIndex]).trim();
        }

        // Append new log entry (separated by comma and newline if existing log exists)
        if (existingLog) {
            row[logUpdateColIndex] = existingLog + ',\n' + logEntry;
        } else {
            row[logUpdateColIndex] = logEntry;
        }

        // Write back to sheet
        sh.getRange(rowNum, 1, 1, row.length).setValues([row]);

        Logger.log('Row updated successfully');

        return {
            success: true
        };
    } catch (e) {
        Logger.log('Error in updateNote: ' + e.toString());
        return {
            success: false,
            error: e.toString()
        };
    }
}

/******************** DELETE NOTE *************************/
function deleteNoteByRow(rowNum) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                success: false,
                error: "Sheet not found"
            };
        }
        sh.deleteRow(rowNum);
        return {
            success: true
        };
    } catch (e) {
        return {
            success: false,
            error: e.toString()
        };
    }
}

/******************** SEARCH NOTES *************************/
function searchNotes(searchTerm) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                headers: [],
                rows: [],
                error: "Sheet not found"
            };
        }

        const v = sh.getDataRange().getValues();
        if (v.length < 2) {
            return {
                headers: v[0] || [],
                rows: []
            };
        }

        const headers = v.shift();
        const searchLower = searchTerm.toLowerCase();

        // Filter rows that contain search term (fuzzy search)
        const filteredRows = [];
        v.forEach((row, originalIdx) => {
            // Search in all columns
            const matches = row.some(cell => {
                const cellStr = String(cell || '').toLowerCase();
                return cellStr.includes(searchLower);
            });
            if (matches) {
                filteredRows.push({
                    row: row,
                    rowNum: originalIdx + 2 // +2 because row 1 is header, data starts at row 2
                });
            }
        });

        // Convert to objects with row numbers
        const rowsWithNum = filteredRows.map(item => {
            const o = {
                _rowNum: item.rowNum
            };
            headers.forEach((h, i) => {
                o[h] = serializeValue(item.row[i]);
            });
            return o;
        });

        // Helper function to serialize values
        function serializeValue(val) {
            if (val === null || val === undefined) return '';
            if (val instanceof Date) {
                // Format date in Vietnam timezone
                return Utilities.formatDate(val, "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss");
            }
            if (Array.isArray(val)) {
                return val.map(serializeValue);
            }
            return val;
        }

        return {
            headers: headers,
            rows: rowsWithNum
        };
    } catch (e) {
        return {
            headers: [],
            rows: [],
            error: e.toString()
        };
    }
}

/******************** FILTER NOTES BY DATE *************************/
function filterNotesByDate(dateFrom, dateTo) {
    try {
        const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
        if (!sh) {
            return {
                headers: [],
                rows: [],
                error: "Sheet not found"
            };
        }

        const v = sh.getDataRange().getValues();
        if (v.length < 2) {
            return {
                headers: v[0] || [],
                rows: []
            };
        }

        const headers = v.shift();
        const timeColIndex = headers.indexOf('Time') !== -1 ? headers.indexOf('Time') : 0;

        // Parse dates (assuming they come as YYYY-MM-DD strings)
        // Parse in Vietnam timezone (GMT+7)
        let fromDate = null;
        let toDate = null;

        if (dateFrom) {
            // Parse date in Vietnam timezone
            const dateParts = dateFrom.split('-');
            fromDate = new Date(Date.UTC(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]), 0, 0, 0));
            // Adjust for Vietnam timezone (UTC+7)
            fromDate.setTime(fromDate.getTime() - (7 * 60 * 60 * 1000));
        }
        if (dateTo) {
            // Parse date in Vietnam timezone
            const dateParts = dateTo.split('-');
            toDate = new Date(Date.UTC(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]), 23, 59, 59));
            // Adjust for Vietnam timezone (UTC+7)
            toDate.setTime(toDate.getTime() - (7 * 60 * 60 * 1000));
            toDate.setMilliseconds(999);
        }

        // Filter rows by date range
        const filteredRows = v.filter(row => {
            const rowDate = row[timeColIndex];
            if (!rowDate) return false;

            // Convert row date to Date object if it's a string
            let rowDateObj = rowDate instanceof Date ? rowDate : new Date(rowDate);

            // Compare dates (both should be in same timezone)
            if (fromDate) {
                const fromTime = fromDate.getTime();
                const rowTime = rowDateObj.getTime();
                if (rowTime < fromTime) return false;
            }
            if (toDate) {
                const toTime = toDate.getTime();
                const rowTime = rowDateObj.getTime();
                if (rowTime > toTime) return false;
            }
            return true;
        });

        // Convert to objects with row numbers
        // Create a map to track original indices
        const rowIndexMap = new Map();
        v.forEach((row, idx) => {
            const rowKey = JSON.stringify(row);
            if (!rowIndexMap.has(rowKey)) {
                rowIndexMap.set(rowKey, idx + 2); // +2 because row 1 is header
            }
        });

        const rowsWithNum = filteredRows.map(row => {
            const rowKey = JSON.stringify(row);
            const originalIndex = rowIndexMap.get(rowKey) || 2;
            const o = {
                _rowNum: originalIndex
            };
            headers.forEach((h, i) => {
                o[h] = serializeValue(row[i]);
            });
            return o;
        });

        // Helper function to serialize values
        function serializeValue(val) {
            if (val === null || val === undefined) return '';
            if (val instanceof Date) {
                // Format date in Vietnam timezone
                return Utilities.formatDate(val, "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss");
            }
            if (Array.isArray(val)) {
                return val.map(serializeValue);
            }
            return val;
        }

        return {
            headers: headers,
            rows: rowsWithNum
        };
    } catch (e) {
        return {
            headers: [],
            rows: [],
            error: e.toString()
        };
    }
}

function deleteNoteRow(rowIndex) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DB_SHEET_NAME);
    sh.deleteRow(rowIndex);
}

/******************** ADMIN CRUD *************************/
function getUsers() {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    let v = sh.getDataRange().getValues();
    return {
        headers: v.shift(),
        rows: v
    };
}
// hàm hỗ trợ insert bố sung
function getUserByIndex(index) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    let row = sh.getRange(index, 1, 1, 10).getValues()[0];

    return {
        user: row[0],
        password: row[1],
        is_admin: row[2],
        canAdd: row[3],
        canEdit: row[4],
        canDelete: row[5],
        trang_chu: row[6],
        app_taive: row[7],
        bo_sung: row[8],
        quan_tri: row[9],
    };
}

function getUserByUsername(username) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    let data = sh.getDataRange().getValues();
    data.shift();

    for (let r of data) {
        if (String(r[0]).trim() === String(username).trim()) {
            return {
                user: r[0],
                password: r[1],
                is_admin: r[2],
                canAdd: r[3],
                canEdit: r[4],
                canDelete: r[5],
                trang_chu: r[6],
                app_taive: r[7],
                bo_sung: r[8],
                quan_tri: r[9],
            };
        }
    }
    return null;
}

function addUser(r) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    sh.appendRow([
        r.user, r.password, r.is_admin,
        r.canAdd, r.canEdit, r.canDelete,
        r.trang_chu, r.app_taive, r.bo_sung, r.quan_tri
    ]);
    return {
        success: true
    };
}

function updateUser(index, r) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    sh.getRange(index, 1, 1, 10).setValues([
        [
            r.user, r.password, r.is_admin,
            r.canAdd, r.canEdit, r.canDelete,
            r.trang_chu, r.app_taive, r.bo_sung, r.quan_tri
        ]
    ]);
    return {
        success: true
    };
}

function deleteUser(index) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    sh.deleteRow(index);
    return {
        success: true
    };
}

// bỏ sung hàm inserUser
function insertUser(r) {
    let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    sh.appendRow([
        r.user, r.password, r.is_admin,
        r.canAdd, r.canEdit, r.canDelete,
        r.trang_chu, r.app_taive, r.bo_sung, r.quan_tri
    ]);
    return {
        success: true
    };
}