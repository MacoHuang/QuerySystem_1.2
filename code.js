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
function showObjectInfo(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var template = HtmlService.createTemplateFromFile('buildingInfo');
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            template.buildingObject = buildingObject;
            console.log(JSON.stringify(buildingObject));
            return template.evaluate().getContent();
        case 'LAND':
            var landTemplate = HtmlService.createTemplateFromFile('landInfo');
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            landTemplate.landObject = landObject;
            console.log(JSON.stringify(landObject));
            return landTemplate.evaluate().getContent();
    }
    return "";
}
function showObjectA4Info(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            // const buildingTemplate = HtmlService.createTemplateFromFile('buildingA4')
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            // buildingTemplate.buildingObject = buildingObject
            // console.log(JSON.stringify(buildingObject))
            return createContract(objectType, buildingObject);
        // return buildingTemplate.evaluate()
        case 'LAND':
            // const landTemplate = HtmlService.createTemplateFromFile('landA4')
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            // landTemplate.landObject = landObject
            // console.log(JSON.stringify(landObject))
            return createContract(objectType, landObject);
        // return landTemplate.evaluate()
    }
    return "";
}
function searchObjectInfo(objectType, sequenceNumberInSheet) {
    // 修正: 透過 DATA_SPREADSHEET 取得 Sheet，而非 getActive()
    var currentSheet = DATA_SPREADSHEET.getSheetByName(objectType);
    var dataRange = currentSheet === null || currentSheet === void 0 ? void 0 : currentSheet.getDataRange();
    var values = dataRange === null || dataRange === void 0 ? void 0 : dataRange.getValues();
    var headers = values === null || values === void 0 ? void 0 : values.shift();
    var row = values === null || values === void 0 ? void 0 : values.find(function (row) {
        return values.indexOf(row) === sequenceNumberInSheet - 1;
    });
    console.log("row:".concat(row));
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
                contractDateFrom: formatDateString(row[LnadHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LnadHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(buildingObject);
        // const template = HtmlService.createTemplateFromFile('buildingInfo')
        // template.buildingObject = buildingObject
        // console.log(JSON.stringify(buildingObject))
        // return template.evaluate().getContent()
        case 'LAND':
            var landObject = {
                createTime: row[LnadHeaders.CREATE_TIME],
                objectNumber: row[LnadHeaders.OBJECT_NUMBER],
                objectName: row[LnadHeaders.OBJECT_NAME],
                contractType: row[LnadHeaders.CONTRACT_TYPE],
                location: row[LnadHeaders.LOCATION],
                landPattern: row[LnadHeaders.LAND_PATTERN],
                landUsage: row[LnadHeaders.LNAD_USAGE],
                landType: row[LnadHeaders.LNAD_TYPE],
                address: row[LnadHeaders.ADDRESS],
                position: row[LnadHeaders.POSITION],
                valuation: row[LnadHeaders.VALUATION],
                landSize: row[LnadHeaders.LAND_SIZE],
                numberOfOwner: row[LnadHeaders.NUMBER_OF_OWNER],
                roadNearby: row[LnadHeaders.ROAD_NEARBY],
                direction: row[LnadHeaders.DIRECTION],
                waterElectricitySupply: row[LnadHeaders.WATER_ELECTRICITY_SUPPLY],
                width: row[LnadHeaders.WIDTH],
                depth: row[LnadHeaders.DEEPTH],
                buildingCoverageRate: row[LnadHeaders.BUILDING_COVERAGE_RATE],
                volumeRate: row[LnadHeaders.VOLUME_RATE],
                memo: row[LnadHeaders.MEMO],
                contactPerson: row[LnadHeaders.CONTACT_PERSON],
                pictureLink: row[LnadHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LnadHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LnadHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(landObject);
        // const landTemplate = HtmlService.createTemplateFromFile('landInfo')
        // landTemplate.landObject = landObject
        // console.log(JSON.stringify(landObject))
        // return landTemplate.evaluate().getContent()
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
var LnadHeaders;
(function (LnadHeaders) {
    LnadHeaders[LnadHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    LnadHeaders[LnadHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    LnadHeaders[LnadHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    LnadHeaders[LnadHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    LnadHeaders[LnadHeaders["LOCATION"] = 4] = "LOCATION";
    LnadHeaders[LnadHeaders["LAND_PATTERN"] = 5] = "LAND_PATTERN";
    LnadHeaders[LnadHeaders["LNAD_USAGE"] = 6] = "LNAD_USAGE";
    LnadHeaders[LnadHeaders["LNAD_TYPE"] = 7] = "LNAD_TYPE";
    LnadHeaders[LnadHeaders["ADDRESS"] = 8] = "ADDRESS";
    LnadHeaders[LnadHeaders["POSITION"] = 9] = "POSITION";
    LnadHeaders[LnadHeaders["VALUATION"] = 10] = "VALUATION";
    LnadHeaders[LnadHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    LnadHeaders[LnadHeaders["NUMBER_OF_OWNER"] = 12] = "NUMBER_OF_OWNER";
    LnadHeaders[LnadHeaders["ROAD_NEARBY"] = 13] = "ROAD_NEARBY";
    LnadHeaders[LnadHeaders["DIRECTION"] = 14] = "DIRECTION";
    LnadHeaders[LnadHeaders["WATER_ELECTRICITY_SUPPLY"] = 15] = "WATER_ELECTRICITY_SUPPLY";
    LnadHeaders[LnadHeaders["WIDTH"] = 16] = "WIDTH";
    LnadHeaders[LnadHeaders["DEEPTH"] = 17] = "DEEPTH";
    LnadHeaders[LnadHeaders["BUILDING_COVERAGE_RATE"] = 18] = "BUILDING_COVERAGE_RATE";
    LnadHeaders[LnadHeaders["VOLUME_RATE"] = 19] = "VOLUME_RATE";
    LnadHeaders[LnadHeaders["MEMO"] = 20] = "MEMO";
    LnadHeaders[LnadHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    LnadHeaders[LnadHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    LnadHeaders[LnadHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    LnadHeaders[LnadHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    LnadHeaders[LnadHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    LnadHeaders[LnadHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(LnadHeaders || (LnadHeaders = {}));

/**
 * 輔助函數：將 GQL 欄位索引 (A, B, C...) 轉換為您的 Headers 列舉索引 (0, 1, 2...)
 * 警告：GQL 始終從 A 欄開始，因此這是一個基於零的索引映射
 */
function getColumnIndex(columnLetter) {
    const charCode = columnLetter.toUpperCase().charCodeAt(0);
    return charCode - 'A'.charCodeAt(0);
}

/**
 * 輔助函數：解析 GQL 傳回的特殊 JSON 格式
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


function searchObjects(contractType, objectType, objectPattern, objectNmae, valuationFrom, valuationTo, landSizeFrom, landSizeTo, roadNearby, roomFrom, roomTo, isHasParkingSpace, buildingAgeFrom, buildingAgeTo, direction, objectWidthFrom, objectWidthTo, contactPerson) {
    
    var listOfSheet = new Array();
    var sheetNames = [];
    
    // 判斷要查詢哪些工作表
    if (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND') {
        sheetNames.push(objectType);
    } else {
        // 如果沒有指定類型，則搜尋目標檔案中的所有工作表 (這裡只考慮 Building 和 Land)
        sheetNames = ['Building', 'Land'];
    }
    
    var extractedData = [];

    // GQL 查詢邏輯取代了原有的 for 迴圈和 filter
    sheetNames.forEach(sheetName => {
        const currentSheet = DATA_SPREADSHEET.getSheetByName(sheetName);
        if (!currentSheet) return;

        // 🚨 重要：您必須手動在這裡填入 Building 和 Land 工作表的 GID
        // GID 可在 Sheet 網址中找到 (例如: .../edit#gid=0)
        const SHEET_GIDS = { 'BUILDING': 'YOUR_BUILDING_GID', 'LAND': 'YOUR_LAND_GID' }; 
        const GID = SHEET_GIDS[sheetName.toUpperCase()];

        if (!GID) return;

        // 構建 GQL 查詢語句
        // 假設 GQL 查詢所有欄位 (A, B, C...)
        let query = 'SELECT * WHERE 1=1'; 
        
        // 為了避免過於複雜，這裡只示範 Valuation 的條件，您需要根據您的 Header 調整欄位字母
        // 假設 Valuation 是 K 欄 (BuildingHeaders.VALUATION=10 -> K 欄)
        const VALUATION_COL = String.fromCharCode('A'.charCodeAt(0) + BuildingHeaders.VALUATION); // K
        
        if (valuationFrom > 0) {
            query += ` AND ${VALUATION_COL} >= ${valuationFrom}`;
        }
        if (valuationTo > 0) {
            query += ` AND ${VALUATION_COL} <= ${valuationTo}`;
        }
        
        // ... (在這裡加入其他所有篩選條件，轉換為 GQL 語法，例如：AND L >= ${landSizeFrom}) ...
        
        // 4. 執行查詢
        const GQL_URL = `https://docs.google.com/spreadsheets/d/${DATA_SHEET_ID}/gviz/tq?tqx=out:json&gid=${GID}`;
        const finalUrl = `${GQL_URL}&tq=${encodeURIComponent(query)}`;

        try {
            const response = UrlFetchApp.fetch(finalUrl).getContentText();
            const { headers, rows } = parseGqlResponse(response);
            
            // 5. 解析資料並轉換為您預期的格式
            const temp = rows.map((row, index) => {
                let data = {};
                let headerMap = sheetName.toUpperCase() === 'BUILDING' ? BuildingHeaders : LnadHeaders;

                data = {
                    objectType: sheetName,
                    // GQL 返回的資料沒有 sequenceNumberInSheet，這裡必須回傳 -1 或其他預設值
                    sequenceNumberInSheet: index + 1, 
                    objectNumber: row[headerMap.OBJECT_NUMBER],
                    objectName: row[headerMap.OBJECT_NAME],
                    valuation: row[headerMap.VALUATION],
                    landSize: row[headerMap.LAND_SIZE],
                    buildingSize: sheetName.toUpperCase() === 'BUILDING' ? row[headerMap.BUILDING_SIZE] : 0,
                    housePattern: sheetName.toUpperCase() === 'BUILDING' ? row[headerMap.HOUSE_PATTERN] : "",
                    position: row[headerMap.POSITION],
                    location: row[headerMap.LOCATION],
                    address: row[headerMap.ADDRESS],
                    pictureLink: row[headerMap.PICTURE_LINK]
                };
                return data;
            });
            extractedData = extractedData.concat(temp);

        } catch (e) {
            console.error(`GQL 查詢錯誤 (${sheetName}): ` + e.toString());
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
            if (!row[LnadHeaders.POSITION]) {
                return false;
            }
            var value = row[LnadHeaders.POSITION].split(',');
            return value != null && value.length == 2 && !isNaN(value[0]) && !isNaN(value[1]);
        })
            .map(function (row) {
            var objectMapData = {
                objectType: 'land',
                objectNumber: row[LnadHeaders.OBJECT_NUMBER],
                objectName: row[LnadHeaders.OBJECT_NAME],
                contractType: row[LnadHeaders.CONTRACT_TYPE],
                location: row[LnadHeaders.LOCATION],
                position: row[LnadHeaders.POSITION],
                valuation: row[LnadHeaders.VALUATION],
                description: row[LnadHeaders.OBJECT_NAME],
                memo: row[LnadHeaders.MEMO],
                contractPerson: row[LnadHeaders.CONTACT_PERSON]
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
            objectNumberColumn = LnadHeaders.OBJECT_NUMBER;
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