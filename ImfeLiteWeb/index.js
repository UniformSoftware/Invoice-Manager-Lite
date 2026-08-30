var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
(function () {
    "use strict";
    var sProgramName = 'Invoice Manager (Lite)';
    var ignoreNames = ['okncompanyname', 'okncompanyaddress', 'okncompanycitystatezip', 'okncompanycontact', 'okndatabasename', 'oknstatus'];
    var bInLoadingEvent = false;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        window.Promise = OfficeExtension.Promise;
        if (document.readyState === "loading") {
            document.addEventListener("DOMContentLoaded", excelMain);
        }
        else {
            excelMain();
        }
    };
    function excelMain() {
        var bFeatureSupported11;
        var element;
        try {
            bFeatureSupported11 = Office.context.requirements.isSetSupported('ExcelApi', "1.1");
            if (!bFeatureSupported11) {
                document.getElementById('tabs').hidden = true;
                document.getElementById('excel2013').hidden = false;
                return;
            }
            document.querySelector("#quick-start").addEventListener('click', onTabChange, this);
            document.querySelector("#commands").addEventListener('click', onTabChange, this);
            document.querySelector("#about").addEventListener('click', onTabChange, this);
            document.getElementById('cmdOpenSampleTemplate').addEventListener('click', cmdOpenSampleTemplate);
            document.getElementById('cmdClearNew').addEventListener("click", cmdClearNew);
            var options = ["toggleInvoiceID", "toggleInvoiceDate", "toggleAutoOpen", "nextInvoiceID", "numberOfDigitsInInvoiceID", "invoiceIDPrefix"];
            for (var _i = 0, options_1 = options; _i < options_1.length; _i++) {
                var option = options_1[_i];
                element = document.getElementById(option);
                if (element != null)
                    element.addEventListener("change", saveOptions);
            }
            element = document.getElementById('gotoCommands');
            element.addEventListener("click", activateCommandsTab);
            element = document.getElementById('cmdOpenHomePage');
            element.addEventListener("click", cmdOpenHomePage);
            LoadOptions();
        }
        catch (ex) {
            showNotification("Loading add-in", ex);
        }
    }
    function LoadOptions() {
        var checkbox;
        var textBox;
        var sOptionValue;
        var bBoolOptionValue;
        try {
            bInLoadingEvent = true;
            sOptionValue = Office.context.document.settings.get('toggleInvoiceDate');
            if (sOptionValue === null) {
                sOptionValue = 'on';
            }
            checkbox = document.getElementById('toggleInvoiceDate');
            if (checkbox != null)
                checkbox.checked = (sOptionValue == 'on' || sOptionValue == 'true' ? true : false);
            sOptionValue = Office.context.document.settings.get('toggleInvoiceID');
            if (sOptionValue === null) {
                sOptionValue = 'on';
            }
            checkbox = document.getElementById('toggleInvoiceID');
            if (checkbox != null)
                checkbox.checked = (sOptionValue == 'on' || sOptionValue == 'true' ? true : false);
            bBoolOptionValue = Boolean(Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument"));
            checkbox = document.getElementById('toggleAutoOpen');
            if (checkbox != null)
                checkbox.checked = bBoolOptionValue;
            sOptionValue = Office.context.document.settings.get('nextInvoiceID');
            if (sOptionValue === null) {
                sOptionValue = '1';
            }
            textBox = document.getElementById('nextInvoiceID');
            if (textBox != null)
                textBox.value = sOptionValue;
            sOptionValue = Office.context.document.settings.get('numberOfDigitsInInvoiceID');
            if (sOptionValue === null) {
                sOptionValue = '4';
            }
            textBox = document.getElementById('numberOfDigitsInInvoiceID');
            if (textBox != null)
                textBox.value = sOptionValue;
            sOptionValue = Office.context.document.settings.get('invoiceIDPrefix');
            if (sOptionValue === null) {
                sOptionValue = 'INV';
            }
            textBox = document.getElementById('invoiceIDPrefix');
            if (textBox != null)
                textBox.value = sOptionValue;
            // The following part does not cause exceptions.
            try {
                var sPreviousPivot;
                sPreviousPivot = Office.context.document.settings.get('activetab');
                if (sPreviousPivot != null) {
                    //var element = document.getElementById(sPreviousPivot);
                    //if (element != null) element.click();
                    var element = document.getElementById('tabs');
                    if (element != null) {
                        element.setAttribute('activeid', sPreviousPivot);
                    }
                }
            }
            catch (_a) { }
        }
        catch (err) {
            showNotification('Loading Options', err);
        }
        finally {
            bInLoadingEvent = false;
        }
    }
    function cmdOpenHomePage() {
        window.open('https://uniformsoftware.com/invoice#lite', '_blank');
    }
    function cmdOpenSampleTemplate() {
        window.open('https://uniformsoftware.com/invoice/lite/invoice-template.xlsx', '_blank');
    }
    function activateCommandsTab() {
        var element = document.getElementById('commands');
        element.click();
    }
    function saveOptions(event) {
        if (bInLoadingEvent)
            return;
        Excel.run(function (ctx) {
            return __awaiter(this, void 0, void 0, function () {
                var thisObject, sOptionName, sOptionValue, parsed, bError, Value, err_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            bError = false;
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 4, , 5]);
                            thisObject = event.target;
                            sOptionName = thisObject.id;
                            if (thisObject.tagName.indexOf('SWITCH') != -1) {
                                sOptionValue = thisObject.checked ? 'true' : 'false';
                            }
                            else {
                                sOptionValue = thisObject.value.trim();
                            }
                            if (sOptionName === 'nextInvoiceID' || sOptionName === 'numberOfDigitsInInvoiceID') {
                                parsed = parseInt(sOptionValue, 10);
                                if (isNaN(parsed)) {
                                    bError = true;
                                }
                                else {
                                    bError = (parsed.toString() !== sOptionValue) || (parsed <= 0);
                                }
                                if (!bError) {
                                    if (sOptionName === 'numberOfDigitsInInvoiceID') {
                                        if (!(parsed >= 1 && parsed <= 9)) {
                                            bError = true;
                                        }
                                    }
                                }
                                if (bError) {
                                    Value = Office.context.document.settings.get(sOptionName);
                                    thisObject.value = Value;
                                    return [2 /*return*/];
                                }
                            }
                            if (sOptionName === 'toggleAutoOpen') {
                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", sOptionValue === 'true' ? true : false);
                            }
                            else {
                                Office.context.document.settings.set(sOptionName, sOptionValue);
                            }
                            // Ensure settings are saved before continuing. saveAsync is callback-based, so
                            // wrap it in a Promise and await completion. This avoids race conditions where
                            // the add-in may close or context may change before the async save completes.
                            return [4 /*yield*/, new Promise(function (resolve, reject) {
                                    try {
                                        Office.context.document.settings.saveAsync(function (asyncResult) {
                                            if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Failed) {
                                                reject(asyncResult.error || 'saveAsync failed');
                                            }
                                            else {
                                                // Some hosts return a simple object with "status" string; treat non-failure as success
                                                resolve(asyncResult);
                                            }
                                        });
                                    }
                                    catch (ex) {
                                        // If the host doesn't support saveAsync or throws synchronously, propagate the error
                                        reject(ex);
                                    }
                                })];
                        case 2:
                            // Ensure settings are saved before continuing. saveAsync is callback-based, so
                            // wrap it in a Promise and await completion. This avoids race conditions where
                            // the add-in may close or context may change before the async save completes.
                            _a.sent();
                            return [4 /*yield*/, ctx.sync()];
                        case 3:
                            _a.sent();
                            return [3 /*break*/, 5];
                        case 4:
                            err_1 = _a.sent();
                            showNotification('Saving options', 'Name: ' + sOptionName + '\n value: ' + sOptionValue + '\n\nError:' + err_1);
                            return [3 /*break*/, 5];
                        case 5: return [2 /*return*/];
                    }
                });
            });
        });
    }
    function cmdClearNew() {
        Excel.run(function (ctx) {
            return __awaiter(this, void 0, void 0, function () {
                var bSheetUnprotected, toggleInvoiceDate, toggleInvoiceID, activeWorksheet, err_2, element, arrNamedRanges, arrNamedRangeWithFormulas, sRangeNameInLowerCase, iInvoiceDateAddress, iInvoiceIDAddress, iRangeToClear, range, iRangeWithFormulas, oRangeWithFormulas, formulas, btn, ex_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            toggleInvoiceDate = true;
                            toggleInvoiceID = true;
                            ;
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 9, , 10]);
                            activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
                            _a.label = 2;
                        case 2:
                            _a.trys.push([2, 4, , 5]);
                            return [4 /*yield*/, unprotectSheet(activeWorksheet, ctx)];
                        case 3:
                            bSheetUnprotected = _a.sent();
                            if (!bSheetUnprotected) {
                                return [2 /*return*/];
                            }
                            return [3 /*break*/, 5];
                        case 4:
                            err_2 = _a.sent();
                            showNotification('Sheet Protection', 'The sheet is protected. Unprotect it by clicking the Excel ribbon command "Review" and then "Unprotect sheet".\n ');
                            return [2 /*return*/];
                        case 5:
                            element = document.getElementById('toggleInvoiceDate');
                            if (element != null)
                                toggleInvoiceDate = element.checked;
                            element = document.getElementById('toggleInvoiceID');
                            if (element != null)
                                toggleInvoiceID = element.checked;
                            arrNamedRanges = new Array();
                            arrNamedRangeWithFormulas = new Array();
                            return [4 /*yield*/, getNamedRanges(ctx, activeWorksheet, arrNamedRanges)];
                        case 6:
                            _a.sent();
                            for (iRangeToClear = 0; iRangeToClear < arrNamedRanges.length; iRangeToClear++) {
                                range = activeWorksheet.getRange(arrNamedRanges[iRangeToClear].name);
                                range.load(['formulas', 'values', 'rowIndex', 'columnIndex', 'address']);
                                arrNamedRangeWithFormulas.push(range);
                                sRangeNameInLowerCase = arrNamedRanges[iRangeToClear].name.toLowerCase();
                                if (sRangeNameInLowerCase === 'okninvoiceid') {
                                    iInvoiceIDAddress = arrNamedRangeWithFormulas.length - 1;
                                }
                                else if (sRangeNameInLowerCase === 'okninvoicedate') {
                                    iInvoiceDateAddress = arrNamedRangeWithFormulas.length - 1;
                                }
                            }
                            return [4 /*yield*/, ctx.sync()];
                        case 7:
                            _a.sent();
                            if (Array.isArray(arrNamedRangeWithFormulas) === false) {
                                throw 'Unable to check range formulas.';
                            }
                            if (arrNamedRangeWithFormulas.length < 1) {
                                throw 'Checking range formulas failed.';
                            }
                            // Clear
                            for (iRangeWithFormulas = 0; iRangeWithFormulas < arrNamedRangeWithFormulas.length; iRangeWithFormulas++) {
                                oRangeWithFormulas = arrNamedRangeWithFormulas[iRangeWithFormulas];
                                formulas = oRangeWithFormulas.formulas[0][0];
                                if (formulas.toString() !== '') {
                                    if (formulas.toString().substring(0, 1) === '=') {
                                        continue;
                                    }
                                }
                                try {
                                    if (toggleInvoiceDate && iRangeWithFormulas === iInvoiceDateAddress) {
                                        updateInvoiceDate(oRangeWithFormulas);
                                    }
                                    else if (toggleInvoiceID && iRangeWithFormulas === iInvoiceIDAddress) {
                                        updateInvoiceID(oRangeWithFormulas);
                                    }
                                    else {
                                        oRangeWithFormulas.values = '';
                                    }
                                }
                                catch (err) {
                                    throw 'Error on updating cell ' + oRangeWithFormulas.address + '\n\n' + err;
                                }
                            }
                            return [4 /*yield*/, ctx.sync()];
                        case 8:
                            _a.sent();
                            try {
                                btn = document.getElementById('cmdClearNew');
                                if (btn != null)
                                    btn.style.setProperty('cursor', 'pointer');
                            }
                            catch (_b) {
                                //This part does not show error.
                            }
                            return [3 /*break*/, 10];
                        case 9:
                            ex_1 = _a.sent();
                            showNotification(sProgramName, ex_1);
                            return [3 /*break*/, 10];
                        case 10: return [2 /*return*/];
                    }
                });
            });
        })
            .catch(errorHandler);
    }
    function getNamedRanges(ctx, activeWorksheet, arrNamedRanges) {
        return __awaiter(this, void 0, void 0, function () {
            var nameditems, bTagFound, activeWorksheetNameLength, sActiveSheetNamePrefix, bIgnoreThisName, i, j;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        activeWorksheet.load("name");
                        nameditems = ctx.workbook.names;
                        nameditems.load(['items', 'name', 'type', 'value', 'visible']);
                        return [4 /*yield*/, ctx.sync()];
                    case 1:
                        _a.sent();
                        bTagFound = false;
                        activeWorksheetNameLength = activeWorksheet.name.length;
                        sActiveSheetNamePrefix = activeWorksheet.name + '!';
                        for (i = 0; i < nameditems.items.length; i++) {
                            if (nameditems.items[i].name.substring(0, 3) !== 'okn') {
                                continue;
                            }
                            if (nameditems.items[i].isNullObject)
                                continue;
                            // API set 1.1
                            if (nameditems.items[i].type !== 'Range')
                                continue;
                            // API set 1.1
                            if (nameditems.items[i].visible !== true) {
                                continue;
                            }
                            bIgnoreThisName = false;
                            for (j = 0; j < ignoreNames.length; j++) {
                                if (ignoreNames[j] === nameditems.items[i].name.toLowerCase()) {
                                    bIgnoreThisName = true;
                                    break;
                                }
                            }
                            if (bIgnoreThisName) {
                                continue;
                            }
                            // If the name refers to the range on the activeworksheet
                            //console.log(nameditems.items[i].value);
                            // $$$$ do we require the activeworksheet to be named as 'Invoice'
                            try {
                                if (nameditems.items[i].value.substring(0, activeWorksheetNameLength + 1) !== sActiveSheetNamePrefix) {
                                    continue;
                                }
                            }
                            catch (_b) {
                                continue;
                            }
                            if (!bTagFound) {
                                if (nameditems.items[i].name.toLowerCase() === 'okninvoiceid') {
                                    bTagFound = true;
                                }
                            }
                            arrNamedRanges.push(nameditems.items[i]);
                            // API set 1.4
                            //if (nameditems.items[i].scope !== 'Workbook') { continue; }
                        }
                        if (!bTagFound) {
                            throw 'The cell named "oknInvoiceID" could not be found. Please make sure you are using a template downloaded from UniformSoftware.com, and the template is modified correctly.';
                        }
                        if (!Array.isArray(arrNamedRanges)) {
                            throw 'No range to clear. Please make sure you are using a template downloaded from UniformSoftware.com, and the template is modified correctly.';
                        }
                        if (arrNamedRanges.length < 1) {
                            throw 'No named range to clear. Please make sure you are using a template downloaded from UniformSoftware.com, and the template is modified correctly.';
                        }
                        return [2 /*return*/];
                }
            });
        });
    }
    function updateInvoiceDate(oRange) {
        oRange.values = (new Date()).toJSON().substring(0, 10);
    }
    function updateInvoiceID(oRange) {
        var sNewInvoiceID;
        var sPrefix = 'INV';
        var iNumberOfDigits = 4;
        var iNextNumber = 1;
        var element;
        element = document.getElementById('invoiceIDPrefix');
        if (element !== null)
            sPrefix = element.value;
        sNewInvoiceID = sPrefix.trim();
        element = document.getElementById('numberOfDigitsInInvoiceID');
        if (element != null)
            iNumberOfDigits = parseInt(element.value);
        if (isNaN(iNumberOfDigits))
            iNumberOfDigits = 4;
        if (iNumberOfDigits < 1)
            iNumberOfDigits = 4;
        if (iNumberOfDigits > 9)
            iNumberOfDigits = 9;
        element = document.getElementById('nextInvoiceID');
        if (element != null)
            iNextNumber = parseInt(element.value);
        if (isNaN(iNextNumber))
            iNextNumber = 1;
        if (iNextNumber < 0)
            iNextNumber = 1;
        sNewInvoiceID = sNewInvoiceID + pad(iNextNumber.toString(), iNumberOfDigits, '0');
        oRange.values = sNewInvoiceID;
        iNextNumber = iNextNumber + 1;
        if (element != null)
            element.value = iNextNumber;
        Office.context.document.settings.set("nextInvoiceID", iNextNumber);
        // Fire-and-forget save with callback; keep behavior synchronous in this function
        Office.context.document.settings.saveAsync(function (asyncResult) { });
    }
    function pad(n, width, z) {
        z = z || '0';
        n = n + '';
        return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
    }
    function unprotectSheet(sheet, ctx) {
        return __awaiter(this, void 0, void 0, function () {
            var bResult, bFeatureSupported, oRange, Values, err_3, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        bFeatureSupported = Office.context.requirements.isSetSupported('ExcelApi', "1.2");
                        if (!!bFeatureSupported) return [3 /*break*/, 6];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        oRange = sheet.getCell(1, 1);
                        oRange.load('values');
                        return [4 /*yield*/, ctx.sync()];
                    case 2:
                        _a.sent();
                        Values = oRange.values;
                        oRange.values = Values;
                        return [4 /*yield*/, ctx.sync()];
                    case 3:
                        _a.sent();
                        return [2 /*return*/, true];
                    case 4:
                        err_3 = _a.sent();
                        throw 'The sheet is protected.\n Please unprotect the sheet by clicking the "Unprotect sheet" command on Excel "Review" ribbon tab.\n\n' + err_3;
                    case 5: return [3 /*break*/, 11];
                    case 6:
                        sheet.load(['protection', 'protection/protected']);
                        return [4 /*yield*/, ctx.sync()];
                    case 7:
                        _a.sent();
                        if (!!sheet.protection.protected) return [3 /*break*/, 8];
                        bResult = true;
                        return [3 /*break*/, 11];
                    case 8:
                        _a.trys.push([8, 10, , 11]);
                        sheet.protection.unprotect();
                        return [4 /*yield*/, ctx.sync()];
                    case 9:
                        _a.sent();
                        bResult = true;
                        return [3 /*break*/, 11];
                    case 10:
                        err_4 = _a.sent();
                        throw 'Error occured on unprotecting the sheet.\n\nIs this sheet protected with a password?\n\nTry to unprotect the sheet manually by clicking the "Unprotect sheet" button on the "Review" ribbon tab.\n\n' + err_4;
                    case 11: return [2 /*return*/, bResult];
                }
            });
        });
    }
    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        //console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        document.getElementById('notification-header').innerHTML = header;
        document.getElementById('notification-body').innerHTML = content;
        var dialogElement = document.getElementById('messageDialog');
        dialogElement.style.display = "block";
    }
    function onTabChange(event) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var sActiveTab, _a;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        _b.trys.push([0, 3, , 4]);
                                        if (bInLoadingEvent)
                                            return [2 /*return*/];
                                        sActiveTab = event.target.id;
                                        Office.context.document.settings.set('activetab', sActiveTab);
                                        return [4 /*yield*/, new Promise(function (resolve, reject) {
                                                try {
                                                    Office.context.document.settings.saveAsync(function (asyncResult) {
                                                        try {
                                                            if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Failed) {
                                                                reject(asyncResult.error || 'saveAsync failed');
                                                            }
                                                            else {
                                                                resolve(asyncResult);
                                                            }
                                                        }
                                                        catch (e) {
                                                            resolve(asyncResult);
                                                        }
                                                    });
                                                }
                                                catch (ex) {
                                                    reject(ex);
                                                }
                                            })];
                                    case 1:
                                        _b.sent();
                                        return [4 /*yield*/, context.sync()];
                                    case 2:
                                        _b.sent();
                                        return [3 /*break*/, 4];
                                    case 3:
                                        _a = _b.sent();
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); }).catch(errorHandler)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
})();
//# sourceMappingURL=index.js.map