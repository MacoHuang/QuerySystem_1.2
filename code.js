// Compiled using undefined undefined (TypeScript 4.9.5)

/*1.GAS web app 執行身份(我),檔案免分享,google sites權限控制
 *2.優化查詢速度,Google Visualization API Query Language (GQL),導入快取機制 (CacheService)
 *3.文件權限"知道連結者"
 */

/** * 🚨 重要：請將此 ID 替換為您實際的 Google Sheet 檔案 ID
 * 此 ID 可以在 Google Sheet 的網址中找到：
 * https://docs.google.com/spreadsheets/d/這個部分就是ID/edit...
 */
const DATA_SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
// 透過 ID 開啟目標 Sheet，確保 Web App 可以正確存取資料
const DATA_SPREADSHEET = SpreadsheetApp.openById(DATA_SHEET_ID);

var serviceUrl = ScriptApp.getService().getUrl();

function getScriptUrl() {
    var url = ScriptApp.getService().getUrl();
    return url;
}
function onEdit(e) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
        if (e.range.getFormula().toUpperCase() == "=MY_OBJECT_NUMBER()") {
            var activeSheet = e.source.getActiveSheet();
            var objectType = activeSheet.getName().toUpperCase();
            e.range.setValue(createObjectNumber(objectType));
        }
    }
    catch (e) {
    }
    finally {
        lock.releaseLock();
    }
}
function doGet(request) {
    var path = request === null || request === void 0 ? void 0 : request.pathInfo;
    switch (path) {
        case 'map':
            var positions = getAllPositions(); // 已導入快取
            var mapTemplate = HtmlService.createTemplateFromFile('objectMap');
            mapTemplate.positions = JSON.stringify(positions);
            return mapTemplate.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
        case 'index':
        default:
            var template = HtmlService.createTemplateFromFile('index');
            template.serviceUrl = serviceUrl;
            return template.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
}
function showObjectInfo(objectType, objectNumber) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var template = HtmlService.createTemplateFromFile('buildingInfo');
            var dataString = searchObjectInfo(objectType, objectNumber);
            var buildingObject = JSON.parse(dataString);
            template.buildingObject = buildingObject;
            console.log(JSON.stringify(buildingObject));
            return template.evaluate().getContent();
        case 'LAND':
            var landTemplate = HtmlService.createTemplateFromFile('landInfo');
            var landDataString = searchObjectInfo(objectType, objectNumber);
            var landObject = JSON.parse(landDataString);
            landTemplate.landObject = landObject;
            console.log(JSON.stringify(landObject));
            return landTemplate.evaluate().getContent();
    }
    return "";
}
function showObjectA4Info(objectType, objectNumber) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var dataString = searchObjectInfo(objectType, objectNumber);
            var buildingObject = JSON.parse(dataString);
            return createContract(objectType, buildingObject);
        case 'LAND':
            var landDataString = searchObjectInfo(objectType, objectNumber);
            var landObject = JSON.parse(landDataString);
            return createContract(objectType, landObject);
    }
    return "";
}
function searchObjectInfo(objectType, objectNumber) {
    const currentSheet = DATA_SPREADSHEET.getSheetByName(objectType);
    if (!currentSheet) return "";

    const GID = currentSheet.getSheetId();
    if (GID === null || GID === undefined) return "";

    const isBuilding = objectType.toUpperCase() === 'BUILDING';
    const Headers = isBuilding ? BuildingHeaders : LandHeaders;

    const objectNumberCol = toGqlCol(Headers.OBJECT_NUMBER);
    const query = `SELECT * WHERE ${objectNumberCol} = '${objectNumber}'`;

    const GQL_URL = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?tqx=out:json&gid=${GID}`;
    const finalUrl = `${GQL_URL}&tq=${encodeURIComponent(query)}`;

    let row;
    try {
        const responseText = UrlFetchApp.fetch(finalUrl, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText();
        const { headers, rows } = parseGqlResponse(responseText);

        if (!rows || rows.length === 0) {
            console.error(`Object not found via GQL: ${objectType} - ${objectNumber}`);
            return "";
        }
        row = rows[0];
    } catch (e) {
        console.error(`GQL Error in searchObjectInfo for ${objectNumber}: ` + e.toString());
        return "";
    }

    if (!row) {
        return "";
    }
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var buildingObject = {
                createTime: row[BuildingHeaders.CREATE_TIME],
                objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                objectName: row[BuildingHeaders.OBJECT_NAME],
                contractType: row[BuildingHeaders.CONTRACT_TYPE],
                location: row[BuildingHeaders.LOCATION],
                buildingType: row[BuildingHeaders.BUILDING_TYPE],
                housePattern: row[BuildingHeaders.HOUSE_PATTERN],
                floor: row[BuildingHeaders.FLOOR],
                address: row[BuildingHeaders.ADDRESS],
                position: row[BuildingHeaders.POSITION],
                valuation: row[BuildingHeaders.VALUATION],
                landSize: row[BuildingHeaders.LAND_SIZE],
                buildingSize: row[BuildingHeaders.BUILDING_SIZE],
                direction: row[BuildingHeaders.DIRECTION],
                vihecleParkingType: row[BuildingHeaders.VIHECLE_PARKING_TYPE],
                vihecleParkingNumber: row[BuildingHeaders.VIHECLE_PARKING_NUMBER],
                waterSupply: row[BuildingHeaders.WATER_SUPPLY],
                roadNearby: row[BuildingHeaders.ROAD_NEARBY],
                width: row[BuildingHeaders.WIDTH],
                buildingAge: row[BuildingHeaders.BUILDING_AGE],
                memo: row[BuildingHeaders.MEMO],
                contactPerson: row[BuildingHeaders.CONTACT_PERSON],
                pictureLink: row[BuildingHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[BuildingHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[BuildingHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(buildingObject);
        case 'LAND':
            var landObject = {
                createTime: row[LandHeaders.CREATE_TIME],
                objectNumber: row[LandHeaders.OBJECT_NUMBER],
                objectName: row[LandHeaders.OBJECT_NAME],
                contractType: row[LandHeaders.CONTRACT_TYPE],
                location: row[LandHeaders.LOCATION],
                landPattern: row[LandHeaders.LAND_PATTERN],
                landUsage: row[LandHeaders.LAND_USAGE],
                landType: row[LandHeaders.LAND_TYPE],
                address: row[LandHeaders.ADDRESS],
                position: row[LandHeaders.POSITION],
                valuation: row[LandHeaders.VALUATION],
                landSize: row[LandHeaders.LAND_SIZE],
                numberOfOwner: row[LandHeaders.NUMBER_OF_OWNER],
                roadNearby: row[LandHeaders.ROAD_NEARBY],
                direction: row[LandHeaders.DIRECTION],
                waterElectricitySupply: row[LandHeaders.WATER_ELECTRICITY_SUPPLY],
                width: row[LandHeaders.WIDTH],
                depth: row[LandHeaders.DEPTH],
                buildingCoverageRate: row[LandHeaders.BUILDING_COVERAGE_RATE],
                volumeRate: row[LandHeaders.VOLUME_RATE],
                memo: row[LandHeaders.MEMO],
                contactPerson: row[LandHeaders.CONTACT_PERSON],
                pictureLink: row[LandHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LandHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LandHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(landObject);
    }
    return "";
}
function formatDateString(date) {
    try {
        return Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd');
    }
    catch (error) {
        return "";
    }
}
function createObjectNumber(objectType) {
    var objectNumberPrefix = '';
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            break;
        default:
    }
    return objectNumberPrefix + (searchLastNumOfNumberedObjects(objectType) + 1);
}
function createContract(objectType, data) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            return createBuildingContract(data);
        case 'LAND':
            return createLandContract(data);
    }
    return "";
}
function createBuildingContract(data) {
    var googleDocId = '1fE0OZZQ00rcYU38vQWCl4h9kE2oJbHmz5uhb_FtP6Gs'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    // var googleDocId = '1fBHyUGHH0-hVNq2fTZVXKVxCJ0UYHjkpdOhM1jefQgI'; // 測試google doc ID
    // var outputFolderId = '1lSczRQ0HEKQrcK8PgfHvqD2kLRmRtsrH'; // 測試google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderBuildingDoc(doc, data);
    return doc.getUrl();
}
function createLandContract(data) {
    var googleDocId = '1MkGlxmbkGtMayj1ZqHd5y9kIwigZ5ky_ZlwRR1h0hH0'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    // var googleDocId = '1noZPLBuWEowiDHni3p-6RoafbOV45BylHkdocQ39p0Y'; // 測試google doc ID
    // var outputFolderId = '1lSczRQ0HEKQrcK8PgfHvqD2kLRmRtsrH'; // 測試google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderLandDoc(doc, data);
    return doc.getUrl();
}
// 先從樣板合約中複製出一個全新的google doc(this.doc)
function createDoc(googleDocId, outputFolderId, fileName) {
    var file = DriveApp.getFileById(googleDocId);
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    var copy = file.makeCopy(fileName, outputFolder);
    var doc = DocumentApp.openById(copy.getId());
    return doc;
}
// **** 修正: 將所有 Unicode 跳脫序列替換為純中文 ****
function renderBuildingDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{編號}}", data.objectNumber);
    body.replaceText("{{案名}}", data.objectName);
    body.replaceText("{{合約類型}}", data.contractType);
    body.replaceText("{{地區}}", data.location);
    body.replaceText("{{形態}}", data.buildingType);
    body.replaceText("{{格局}}", data.housePattern);
    body.replaceText("{{樓層}}", data.floor.toString());
    body.replaceText("{{地址}}", data.address);
    body.replaceText("{{位置}}", data.position);
    body.replaceText("{{總價}}", data.valuation.toString());
    body.replaceText("{{地坪}}", data.landSize.toString());
    body.replaceText("{{建坪}}", data.buildingSize.toString());
    body.replaceText("{{座向}}", data.direction);
    body.replaceText("{{車位}}", data.vihecleParkingType);
    body.replaceText("{{車位號碼}}", data.vihecleParkingNumber.toString());
    body.replaceText("{{水電}}", data.waterSupply);
    body.replaceText("{{臨路}}", data.roadNearby);
    body.replaceText("{{面寬}}", data.width.toString());
    body.replaceText("{{完成日}}", data.buildingAge);
    body.replaceText("{{備註}}", data.memo);
    body.replaceText("{{聯絡人}}", data.contactPerson);
    body.replaceText("{{圖片連結}}", data.pictureLink);
    body.replaceText("{{合約開始日期}}", data.contractDateFrom);
    body.replaceText("{{合約結束日期}}", data.contractDateTo);
    doc.saveAndClose();
}
// **** 修正: 將所有 Unicode 跳脫序列替換為純中文 ****
function renderLandDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{編號}}", data.objectNumber);
    body.replaceText("{{案名}}", data.objectName);
    body.replaceText("{{合約類型}}", data.contractType);
    body.replaceText("{{地區}}", data.location);
    body.replaceText("{{類別}}", data.landType);
    body.replaceText("{{分區}}", data.landUsage);
    body.replaceText("{{形態}}", data.landPattern);
    body.replaceText("{{地址}}", data.address);
    body.replaceText("{{位置}}", data.position);
    body.replaceText("{{總價}}", data.valuation.toString());
    body.replaceText("{{地坪_1}}", data.landSize.toString());
    body.replaceText("{{地坪_2}}", (Math.round((data.landSize / 293.4) * 100) / 100).toString());
    body.replaceText("{{所有權人數}}", data.numberOfOwner.toString());
    body.replaceText("{{臨路}}", data.roadNearby);
    body.replaceText("{{座向}}", data.direction);
    body.replaceText("{{水電}}", data.waterElectricitySupply);
    body.replaceText("{{面寬}}", data.width.toString());
    body.replaceText("{{縱深}}", data.depth.toString());
    body.replaceText("{{建蔽率}}", data.buildingCoverageRate.toString());
    body.replaceText("{{容積率}}", data.volumeRate.toString());
    body.replaceText("{{備註}}", data.memo);
    body.replaceText("{{聯絡人}}", data.contactPerson);
    body.replaceText("{{圖片連結}}", data.pictureLink);
    body.replaceText("{{合約開始日期}}", data.contractDateFrom);
    body.replaceText("{{合約結束日期}}", data.contractDateTo);
    doc.saveAndClose();
}
var BuildingHeaders;
(function (BuildingHeaders) {
    BuildingHeaders[BuildingHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    BuildingHeaders[BuildingHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    BuildingHeaders[BuildingHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    BuildingHeaders[BuildingHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    BuildingHeaders[BuildingHeaders["LOCATION"] = 4] = "LOCATION";
    BuildingHeaders[BuildingHeaders["BUILDING_TYPE"] = 5] = "BUILDING_TYPE";
    BuildingHeaders[BuildingHeaders["HOUSE_PATTERN"] = 6] = "HOUSE_PATTERN";
    BuildingHeaders[BuildingHeaders["FLOOR"] = 7] = "FLOOR";
    BuildingHeaders[BuildingHeaders["ADDRESS"] = 8] = "ADDRESS";
    BuildingHeaders[BuildingHeaders["POSITION"] = 9] = "POSITION";
    BuildingHeaders[BuildingHeaders["VALUATION"] = 10] = "VALUATION";
    BuildingHeaders[BuildingHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    BuildingHeaders[BuildingHeaders["BUILDING_SIZE"] = 12] = "BUILDING_SIZE";
    BuildingHeaders[BuildingHeaders["DIRECTION"] = 13] = "DIRECTION";
    BuildingHeaders[BuildingHeaders["VIHECLE_PARKING_TYPE"] = 14] = "VIHECLE_PARKING_TYPE";
    BuildingHeaders[BuildingHeaders["VIHECLE_PARKING_NUMBER"] = 15] = "VIHECLE_PARKING_NUMBER";
    BuildingHeaders[BuildingHeaders["WATER_SUPPLY"] = 16] = "WATER_SUPPLY";
    BuildingHeaders[BuildingHeaders["ROAD_NEARBY"] = 17] = "ROAD_NEARBY";
    BuildingHeaders[BuildingHeaders["WIDTH"] = 18] = "WIDTH";
    BuildingHeaders[BuildingHeaders["BUILDING_AGE"] = 19] = "BUILDING_AGE";
    BuildingHeaders[BuildingHeaders["MEMO"] = 20] = "MEMO";
    BuildingHeaders[BuildingHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    BuildingHeaders[BuildingHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    BuildingHeaders[BuildingHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    BuildingHeaders[BuildingHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(BuildingHeaders || (BuildingHeaders = {}));
var LandHeaders;
(function (LandHeaders) {
    LandHeaders[LandHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    LandHeaders[LandHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    LandHeaders[LandHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    LandHeaders[LandHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    LandHeaders[LandHeaders["LOCATION"] = 4] = "LOCATION";
    LandHeaders[LandHeaders["LAND_PATTERN"] = 5] = "LAND_PATTERN";
    LandHeaders[LandHeaders["LAND_USAGE"] = 6] = "LAND_USAGE";
    LandHeaders[LandHeaders["LAND_TYPE"] = 7] = "LAND_TYPE";
    LandHeaders[LandHeaders["ADDRESS"] = 8] = "ADDRESS";
    LandHeaders[LandHeaders["POSITION"] = 9] = "POSITION";
    LandHeaders[LandHeaders["VALUATION"] = 10] = "VALUATION";
    LandHeaders[LandHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    LandHeaders[LandHeaders["NUMBER_OF_OWNER"] = 12] = "NUMBER_OF_OWNER";
    LandHeaders[LandHeaders["ROAD_NEARBY"] = 13] = "ROAD_NEARBY";
    LandHeaders[LandHeaders["DIRECTION"] = 14] = "DIRECTION";
    LandHeaders[LandHeaders["WATER_ELECTRICITY_SUPPLY"] = 15] = "WATER_ELECTRICITY_SUPPLY";
    LandHeaders[LandHeaders["WIDTH"] = 16] = "WIDTH";
    LandHeaders[LandHeaders["DEPTH"] = 17] = "DEPTH";
    LandHeaders[LandHeaders["BUILDING_COVERAGE_RATE"] = 18] = "BUILDING_COVERAGE_RATE";
    LandHeaders[LandHeaders["VOLUME_RATE"] = 19] = "VOLUME_RATE";
    LandHeaders[LandHeaders["MEMO"] = 20] = "MEMO";
    LandHeaders[LandHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    LandHeaders[LandHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    LandHeaders[LandHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    LandHeaders[LandHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    LandHeaders[LandHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    LandHeaders[LandHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(LandHeaders || (LandHeaders = {}));

/**
 * 輔助函數：將 GQL 傳回的特殊 JSON 格式
 */
function parseGqlResponse(response) {
    const json = response.substring(response.indexOf('{'), response.lastIndexOf('}') + 1);
    const data = JSON.parse(json);
    if (!data.table || !data.table.rows) {
        return [];
    }
    
    const headers = data.table.cols.map(col => col.label);
    const rows = data.table.rows.map(row => {
        const values = row.c.map(cell => (cell && cell.v !== undefined) ? cell.v : null);
        return values;
    });
    
    return { headers, rows };
}

/**
 * 輔助函數：將基於 0 的欄位索引轉換為 GQL 的欄位字母 (例如 0 -> A, 1 -> B)
 */
function toGqlCol(index) {
    return String.fromCharCode('A'.charCodeAt(0) + index);
}

function searchObjects(contractType, objectType, objectPattern, objectName, valuationFrom, valuationTo, landSizeFrom, landSizeTo, roadNearby, roomFrom, roomTo, isHasParkingSpace, buildingAgeFrom, buildingAgeTo, direction, objectWidthFrom, objectWidthTo, contactPerson) {
    
    var sheetNames = [];
    
    // 判斷要查詢哪些工作表
    if (objectType && (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND')) {
        sheetNames.push(objectType);
    } else {
        // 如果沒有指定類型，則搜尋全部
        sheetNames = ['Building', 'Land'];
    }
    
    var extractedData = [];

    sheetNames.forEach(sheetName => {
        const currentSheet = DATA_SPREADSHEET.getSheetByName(sheetName);
        if (!currentSheet) return;

        const GID = currentSheet.getSheetId(); // 動態取得 GID
        if (GID === null || GID === undefined) return;

        const isBuilding = sheetName.toUpperCase() === 'BUILDING';
        const Headers = isBuilding ? BuildingHeaders : LandHeaders;
        let queryConditions = [];
        
        // --- 動態建立 GQL WHERE 查詢條件 ---
        if (contractType) {
            queryConditions.push(`${toGqlCol(Headers.CONTRACT_TYPE)} = '${contractType}'`);
        }

        if (objectName) {
            const keywords = objectName.split(' ').filter(k => k);
            keywords.forEach(keyword => {
                queryConditions.push(`(${toGqlCol(Headers.OBJECT_NAME)} like '%${keyword}%' or ${toGqlCol(Headers.ADDRESS)} like '%${keyword}%' or ${toGqlCol(Headers.LOCATION)} like '%${keyword}%')`);
            });
        }
        
        if (valuationFrom > 0) queryConditions.push(`${toGqlCol(Headers.VALUATION)} >= ${valuationFrom}`);
        if (valuationTo > 0) queryConditions.push(`${toGqlCol(Headers.VALUATION)} <= ${valuationTo}`);

        if (landSizeFrom > 0) queryConditions.push(`${toGqlCol(Headers.LAND_SIZE)} >= ${landSizeFrom}`);
        if (landSizeTo > 0) queryConditions.push(`${toGqlCol(Headers.LAND_SIZE)} <= ${landSizeTo}`);

        if (roadNearby) {
            const [min, max] = roadNearby.split('|');
            queryConditions.push(`(${toGqlCol(Headers.ROAD_NEARBY)} >= ${min} and ${toGqlCol(Headers.ROAD_NEARBY)} <= ${max})`);
        }

        if (objectWidthFrom > 0) queryConditions.push(`${toGqlCol(Headers.WIDTH)} >= ${objectWidthFrom}`);
        if (objectWidthTo > 0) queryConditions.push(`${toGqlCol(Headers.WIDTH)} <= ${objectWidthTo}`);

        if (direction) {
            queryConditions.push(`${toGqlCol(Headers.DIRECTION)} = '${direction}'`);
        }

        if (contactPerson) {
            queryConditions.push(`${toGqlCol(Headers.CONTACT_PERSON)} like '%${contactPerson}%'`);
        }

        if (objectPattern && objectPattern.length > 0) {
            const patternCol = isBuilding ? toGqlCol(BuildingHeaders.BUILDING_TYPE) : toGqlCol(LandHeaders.LAND_PATTERN);
            const patternConditions = objectPattern.map(p => `${patternCol} = '${p}'`);
            if (patternConditions.length > 0) {
                queryConditions.push(`(${patternConditions.join(' OR ')})`);
            }
        }

        if (isBuilding) {
            if (isHasParkingSpace === '1') { // 有車位
                queryConditions.push(`${toGqlCol(BuildingHeaders.VIHECLE_PARKING_TYPE)} is not null`);
            } else if (isHasParkingSpace === '0') { // 無車位
                queryConditions.push(`(${toGqlCol(BuildingHeaders.VIHECLE_PARKING_TYPE)} is null or ${toGqlCol(BuildingHeaders.VIHECLE_PARKING_TYPE)} = '')`);
            }

            // 建物完成年 (假設 BUILDING_AGE 欄位是年份或可按字串比較的日期)
            if (buildingAgeFrom) queryConditions.push(`${toGqlCol(BuildingHeaders.BUILDING_AGE)} >= '${buildingAgeFrom}'`);
            if (buildingAgeTo) queryConditions.push(`${toGqlCol(BuildingHeaders.BUILDING_AGE)} <= '${buildingAgeTo}'`);

            // 房數 (格局) - 處理文字格式如 "3房2廳"
            const roomCol = toGqlCol(BuildingHeaders.HOUSE_PATTERN);
            let roomConditions = [];
            if (roomFrom > 0 && roomTo > 0 && Number(roomTo) >= Number(roomFrom)) {
                for (let i = Number(roomFrom); i <= Number(roomTo); i++) {
                    roomConditions.push(`${roomCol} like '${i}房%'`);
                }
            } else if (roomFrom > 0) {
                roomConditions.push(`${roomCol} like '${Number(roomFrom)}房%'`);
            } else if (roomTo > 0) {
                for (let i = 1; i <= Number(roomTo); i++) {
                    roomConditions.push(`${roomCol} like '${i}房%'`);
                }
            }
            if(roomConditions.length > 0) {
                queryConditions.push(`(${roomConditions.join(' OR ')})`);
            }
        }
        
        let query = 'SELECT *';
        if (queryConditions.length > 0) {
            query += ' WHERE ' + queryConditions.join(' AND ');
        }
        
        const GQL_URL = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?tqx=out:json&gid=${GID}`;
        const finalUrl = `${GQL_URL}&tq=${encodeURIComponent(query)}`;

        try {
            const response = UrlFetchApp.fetch(finalUrl, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText();
            const { headers, rows } = parseGqlResponse(response);
            
            const temp = rows.map((row) => {
                let headerMap = isBuilding ? BuildingHeaders : LandHeaders;
                // 回傳的資料結構
                return {
                    objectType: sheetName,
                    sequenceNumberInSheet: row[headerMap.OBJECT_NUMBER], // **重要**: 傳遞 objectNumber 以供後續查詢
                    objectNumber: row[headerMap.OBJECT_NUMBER],
                    objectName: row[headerMap.OBJECT_NAME],
                    valuation: row[headerMap.VALUATION],
                    landSize: row[headerMap.LAND_SIZE],
                    buildingSize: isBuilding ? row[headerMap.BUILDING_SIZE] : 0,
                    housePattern: isBuilding ? row[headerMap.HOUSE_PATTERN] : "",
                    position: row[headerMap.POSITION],
                    location: row[headerMap.LOCATION],
                    address: row[headerMap.ADDRESS],
                    pictureLink: row[headerMap.PICTURE_LINK]
                };
            });
            extractedData = extractedData.concat(temp);

        } catch (e) {
            console.error(`GQL 查詢錯誤 (${sheetName}) for query "${query}": ` + e.toString());
        }
    });

    console.log("extractedData.length:".concat(extractedData.length));
    return JSON.stringify(extractedData);
}

function getAllPositions() {
    // 導入 CacheService
    const CACHE_KEY = 'all_object_positions';
    const CACHE_EXPIRATION_SECONDS = 900; // 15 分鐘
    const cache = CacheService.getScriptCache();
    const cachedPositions = cache.get(CACHE_KEY);

    if (cachedPositions) {
        console.log("Returning positions from cache.");
        return JSON.parse(cachedPositions);
    }
    
    // 執行原有邏輯 (讀取資料)
    var buildingSheet = DATA_SPREADSHEET.getSheetByName('Building');
    var landSheet = DATA_SPREADSHEET.getSheetByName('Land');
    var buildingDataRange = buildingSheet === null || buildingSheet === void 0 ? void 0 : buildingSheet.getDataRange();
    var landDataRange = landSheet === null || landSheet === void 0 ? void 0 : landSheet.getDataRange();
    var buildingValues = buildingDataRange === null || buildingDataRange === void 0 ? void 0 : buildingDataRange.getValues();
    var landValues = landDataRange === null || landDataRange === void 0 ? void 0 : landDataRange.getValues();
    var buildingHeaders = buildingValues === null || buildingValues === void 0 ? void 0 : buildingValues.shift();
    var landHeaders = landValues === null || landValues === void 0 ? void 0 : landValues.shift();
    var positions = new Array();
    if (buildingHeaders && buildingValues) {
        positions = positions.concat(buildingValues
            .filter(function (row) {
            if (!row[BuildingHeaders.POSITION]) {
                return false;
            }
            var value = row[BuildingHeaders.POSITION].split(' ')[0];
            return value !== '' && value !== null && value !== undefined && isNaN(Number(value));
        })
            .map(function (row) {
            var objectMapData = {
                objectType: 'building',
                objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                objectName: row[BuildingHeaders.OBJECT_NAME],
                contractType: row[BuildingHeaders.CONTRACT_TYPE],
                location: row[BuildingHeaders.LOCATION],
                position: row[BuildingHeaders.POSITION].split(' ')[0],
                valuation: row[BuildingHeaders.VALUATION],
                description: row[BuildingHeaders.OBJECT_NAME],
                memo: row[BuildingHeaders.MEMO],
                contractPerson: row[BuildingHeaders.CONTACT_PERSON]
            };
            return objectMapData;
        }));
    }
    if (landHeaders && landValues) {
        positions = positions.concat(landValues
            .filter(function (row) {
            if (!row[LandHeaders.POSITION]) {
                return false;
            }
            var value = row[LandHeaders.POSITION].split(',');
            return value != null && value.length == 2 && !isNaN(value[0]) && !isNaN(value[1]);
        })
            .map(function (row) {
            var objectMapData = {
                objectType: 'land',
                objectNumber: row[LandHeaders.OBJECT_NUMBER],
                objectName: row[LandHeaders.OBJECT_NAME],
                contractType: row[LandHeaders.CONTRACT_TYPE],
                location: row[LandHeaders.LOCATION],
                position: row[LandHeaders.POSITION],
                valuation: row[LandHeaders.VALUATION],
                description: row[LandHeaders.OBJECT_NAME],
                memo: row[LandHeaders.MEMO],
                contractPerson: row[LandHeaders.CONTACT_PERSON]
            };
            return objectMapData;
        }));
    }
    
    // 將結果存入快取
    cache.put(CACHE_KEY, JSON.stringify(positions), CACHE_EXPIRATION_SECONDS);
    
    return positions;
}
var BuildingObjectData = /** @class */ (function () {
    function BuildingObjectData() {
    }
    return BuildingObjectData;
}());
var LandObjectData = /** @class */ (function () {
    function LandObjectData() {
    }
    return LandObjectData;
}());
var ObjectMapData = /** @class */ (function () {
    function ObjectMapData() {
    }
    return ObjectMapData;
}());
function searchLastNumOfNumberedObjects(objectType) {
    var listOfSheet = new Array();
    
    // 修正: 透過 DATA_SPREADSHEET 取得目標 Sheet，而非 getActive()
    if (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND') {
        var currentSheet = DATA_SPREADSHEET.getSheetByName(objectType);
        if (currentSheet) {
            listOfSheet.push(currentSheet);
        }
    } else {
        // 取得目標檔案中的所有工作表
        listOfSheet = DATA_SPREADSHEET.getSheets();
    }

    // if object type is building, then the rule of object number is a 'A' + last number of numbered objects plus 1
    // if object type is land, then the rule of object number is a 'B' + last number of numbered objects plus 1
    var objectNumberPrefix = '';
    var objectNumberColumn = 0;
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            objectNumberColumn = BuildingHeaders.OBJECT_NUMBER;
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            objectNumberColumn = LandHeaders.OBJECT_NUMBER;
            break;
        default:
    }
    var lastNumberOfObjectNumber = '';
    for (var _i = 0, listOfSheet_2 = listOfSheet; _i < listOfSheet_2.length; _i++) {
        var currentSheet = listOfSheet_2[_i];
        var dataRange = currentSheet.getDataRange();
        var values = dataRange.getValues();
        var headers = values.shift();
        var objectNumbers = values.map(function (row) {
            return row[objectNumberColumn];
        });
        lastNumberOfObjectNumber = objectNumbers.reduce(function (prev, current) {
            var isHasPrefix = current.toString().startsWith(objectNumberPrefix);
            var currentNumberPart = Number(current.toString().substring(1));
            var prevNumberPart = Number(prev.toString().substring(1));
            var isCurrentANumber = !isNaN(currentNumberPart);
            var isPrevANumber = !isNaN(prevNumberPart);
            if (!isPrevANumber) {
                prevNumberPart = 0;
            }
            if (isHasPrefix && isCurrentANumber) {
                return currentNumberPart > prevNumberPart ? current : prev;
            }
            return prev;
        });
        // const numberedObjectNumbers = objectNumbers.filter(function(objectNumber) {
        //     const isHasPrefix = objectNumber.toString().startsWith(objectNumberPrefix)
        //     const isNumber = !isNaN(Number(objectNumber.toString().substr(1)))
        //     return isHasPrefix && isNumber
        // })
        // numOfNumberedObjects += numberedObjectNumbers.length
    }
    return Number(lastNumberOfObjectNumber.toString().substring(1));
}