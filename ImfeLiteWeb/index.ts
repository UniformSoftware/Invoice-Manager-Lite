(function () {
    "use strict";

    const sProgramName: string = 'Invoice Manager (Lite)'
    const ignoreNames = ['okncompanyname', 'okncompanyaddress', 'okncompanycitystatezip', 'okncompanycontact', 'okndatabasename', 'oknstatus'];
    var bInLoadingEvent = false;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {

        (window as any).Promise = OfficeExtension.Promise;

        if (document.readyState === "loading") {
            document.addEventListener("DOMContentLoaded", excelMain);
        }
        else {
            excelMain();
        }
    };

    function excelMain() {

        var bFeatureSupported11: boolean
        var element;

        try {
            bFeatureSupported11 = Office.context.requirements.isSetSupported('ExcelApi', "1.1");
            if (!bFeatureSupported11) {
                document.getElementById('tabs').hidden = true;
                document.getElementById('excel2013').hidden = false;
                return;
            }

            document.querySelector("#quick-start").addEventListener('click', onTabChange,this);
            document.querySelector("#commands").addEventListener('click', onTabChange, this);
            document.querySelector("#about").addEventListener('click', onTabChange, this);

            document.getElementById('cmdOpenSampleTemplate').addEventListener('click', cmdOpenSampleTemplate);
            document.getElementById('cmdClearNew').addEventListener("click", cmdClearNew);
            const options = ["toggleInvoiceID", "toggleInvoiceDate", "toggleAutoOpen", "nextInvoiceID", "numberOfDigitsInInvoiceID", "invoiceIDPrefix"];
            for (const option of options) {
                element = document.getElementById(option);
                if (element != null) element.addEventListener("change", saveOptions);
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
        let bBoolOptionValue: boolean;

        try {

            bInLoadingEvent = true;

            sOptionValue = Office.context.document.settings.get('toggleInvoiceDate');
            if (sOptionValue === null) { sOptionValue = 'on'; }
            checkbox = document.getElementById('toggleInvoiceDate') as HTMLInputElement | null;
            if (checkbox != null) checkbox.checked = (sOptionValue == 'on' || sOptionValue == 'true' ? true : false);

            sOptionValue = Office.context.document.settings.get('toggleInvoiceID');
            if (sOptionValue === null) { sOptionValue = 'on'; }
            checkbox = document.getElementById('toggleInvoiceID') as HTMLInputElement | null;
            if (checkbox != null) checkbox.checked = (sOptionValue == 'on' || sOptionValue == 'true' ? true : false);

            bBoolOptionValue = Boolean(Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument"));
            checkbox = document.getElementById('toggleAutoOpen') as HTMLInputElement | null;
            if (checkbox != null) checkbox.checked = bBoolOptionValue;

            sOptionValue = Office.context.document.settings.get('nextInvoiceID');
            if (sOptionValue === null) { sOptionValue = '1'; }
            textBox = document.getElementById('nextInvoiceID') as HTMLInputElement | null;
            if (textBox != null) textBox.value = sOptionValue;


            sOptionValue = Office.context.document.settings.get('numberOfDigitsInInvoiceID');
            if (sOptionValue === null) { sOptionValue = '4'; }
            textBox = document.getElementById('numberOfDigitsInInvoiceID') as HTMLInputElement | null;
            if (textBox != null) textBox.value = sOptionValue;


            sOptionValue = Office.context.document.settings.get('invoiceIDPrefix');
            if (sOptionValue === null) { sOptionValue = 'INV'; }
            textBox = document.getElementById('invoiceIDPrefix') as HTMLInputElement | null;
            if (textBox != null) textBox.value = sOptionValue;

            // The following part does not cause exceptions.
            try {
                var sPreviousPivot: string;
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

            catch { }
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

        if (bInLoadingEvent) return;
        Excel.run(async function (ctx) {
            var thisObject;
            var sOptionName: string;
            var sOptionValue: string;
            var parsed;
            var bError: boolean = false;

            try {
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
                    if (isNaN(parsed)) { bError = true; }
                    else {
                        bError = (parsed.toString() !== sOptionValue) || (parsed <= 0);
                    }
                    if (!bError) {
                        if (sOptionName === 'numberOfDigitsInInvoiceID') {
                            if (!(parsed >= 1 && parsed <= 9)) { bError = true; }
                        }
                    }

                    if (bError) {
                        var Value = Office.context.document.settings.get(sOptionName);
                        thisObject.value = Value;
                        return;
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
                await new Promise((resolve, reject) => {
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
                });
                await ctx.sync();
            }
            catch (err) {
                showNotification('Saving options', 'Name: ' + sOptionName + '\n value: ' + sOptionValue + '\n\nError:' + err);
            }
        });
    }
    function cmdClearNew() {
        Excel.run(async function (ctx) {
            var bSheetUnprotected: boolean;

            var toggleInvoiceDate: boolean = true;
            var toggleInvoiceID: boolean = true;;

            try {

                // Unprotect sheet if necessary

                var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();

                try {
                    bSheetUnprotected = await unprotectSheet(activeWorksheet, ctx);
                    if (!bSheetUnprotected) { return; }
                }
                catch (err) {
                    showNotification('Sheet Protection', 'The sheet is protected. Unprotect it by clicking the Excel ribbon command "Review" and then "Unprotect sheet".\n ');
                    return;
                }

                // Options

                var element;
                element = document.getElementById('toggleInvoiceDate');
                if (element != null) toggleInvoiceDate = element.checked;
                element = document.getElementById('toggleInvoiceID');
                if (element != null) toggleInvoiceID = element.checked;

                // Load named ranges and the formulas.

                var arrNamedRanges = new Array();
                var arrNamedRangeWithFormulas = new Array();

                await getNamedRanges(ctx, activeWorksheet, arrNamedRanges);

                var sRangeNameInLowerCase;
                var iInvoiceDateAddress;
                var iInvoiceIDAddress;

                for (var iRangeToClear = 0; iRangeToClear < arrNamedRanges.length; iRangeToClear++) {
                    var range = activeWorksheet.getRange(arrNamedRanges[iRangeToClear].name);
                    range.load(['formulas', 'values', 'rowIndex', 'columnIndex', 'address']);
                    arrNamedRangeWithFormulas.push(range);
                    sRangeNameInLowerCase = arrNamedRanges[iRangeToClear].name.toLowerCase();
                    if (sRangeNameInLowerCase === 'okninvoiceid') {
                        iInvoiceIDAddress = arrNamedRangeWithFormulas.length - 1
                    }
                    else if (sRangeNameInLowerCase === 'okninvoicedate') {
                        iInvoiceDateAddress = arrNamedRangeWithFormulas.length - 1
                    }
                }
                await ctx.sync();

                if (Array.isArray(arrNamedRangeWithFormulas) === false) {
                    throw 'Unable to check range formulas.';
                }
                if (arrNamedRangeWithFormulas.length < 1) {
                    throw 'Checking range formulas failed.';
                }

                // Clear

                for (var iRangeWithFormulas = 0; iRangeWithFormulas < arrNamedRangeWithFormulas.length; iRangeWithFormulas++) {
                    var oRangeWithFormulas = arrNamedRangeWithFormulas[iRangeWithFormulas];
                    var formulas = oRangeWithFormulas.formulas[0][0];

                    if (formulas.toString() !== '') {
                        if (formulas.toString().substring(0, 1) === '=') { continue; }
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

                await ctx.sync();
                try {
                    var btn
                    btn = document.getElementById('cmdClearNew');
                    if (btn != null) btn.style.setProperty('cursor', 'pointer');
                }
                catch {
                    //This part does not show error.
                }
            }
            catch (ex) {
                showNotification(sProgramName, ex);
            }
        })
            .catch(errorHandler);
    }
    async function getNamedRanges(ctx, activeWorksheet, arrNamedRanges: object[]) {
       activeWorksheet.load("name");

       var nameditems = ctx.workbook.names;
       nameditems.load(['items', 'name', 'type', 'value', 'visible']);
       await ctx.sync();

       var bTagFound = false;
       var activeWorksheetNameLength = activeWorksheet.name.length;
       var sActiveSheetNamePrefix = activeWorksheet.name + '!'
       var bIgnoreThisName;

       for (var i = 0; i < nameditems.items.length; i++) {

           if (nameditems.items[i].name.substring(0, 3) !== 'okn') { continue; }
           if (nameditems.items[i].isNullObject) continue;

           // API set 1.1
           if (nameditems.items[i].type !== 'Range') continue;

           // API set 1.1
           if (nameditems.items[i].visible !== true) { continue; }

           bIgnoreThisName = false;
           for (var j = 0; j < ignoreNames.length; j++) {
               if (ignoreNames[j] === nameditems.items[i].name.toLowerCase()) {
                   bIgnoreThisName = true;
                   break;
               }
           }
           if (bIgnoreThisName) { continue; }

           // If the name refers to the range on the activeworksheet
           //console.log(nameditems.items[i].value);
           // $$$$ do we require the activeworksheet to be named as 'Invoice'
           try {
               if (nameditems.items[i].value.substring(0, activeWorksheetNameLength + 1) !== sActiveSheetNamePrefix) { continue; }
           }
           catch {
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
    }
    function updateInvoiceDate(oRange) {
        oRange.values = (new Date()).toJSON().substring(0, 10);
    }
    function updateInvoiceID(oRange) {
        var sNewInvoiceID: string;
        var sPrefix: string = 'INV';
        var iNumberOfDigits = 4;
        var iNextNumber = 1;
        var element;

        element = document.getElementById('invoiceIDPrefix');
        if (element !== null) sPrefix = element.value;
        sNewInvoiceID = sPrefix.trim();

        element = document.getElementById('numberOfDigitsInInvoiceID')
        if (element != null) iNumberOfDigits = parseInt(element.value);
        if (isNaN(iNumberOfDigits)) iNumberOfDigits = 4;
        if (iNumberOfDigits < 1) iNumberOfDigits = 4;
        if (iNumberOfDigits > 9) iNumberOfDigits = 9;

        element = document.getElementById('nextInvoiceID');
        if (element != null) iNextNumber = parseInt(element.value);
        if (isNaN(iNextNumber)) iNextNumber = 1;
        if (iNextNumber < 0) iNextNumber = 1;

        sNewInvoiceID = sNewInvoiceID + pad(iNextNumber.toString(), iNumberOfDigits, '0');
        oRange.values = sNewInvoiceID;
        iNextNumber = iNextNumber + 1;
        if (element != null) element.value = iNextNumber;
        Office.context.document.settings.set("nextInvoiceID", iNextNumber);
        // Fire-and-forget save with callback; keep behavior synchronous in this function
        Office.context.document.settings.saveAsync(function (asyncResult) { });
    }
    function pad(n, width, z) {
        z = z || '0';
        n = n + '';
        return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
    }
    async function unprotectSheet(sheet: Excel.Worksheet, ctx: Excel.RequestContext) {
        var bResult: boolean;
        var bFeatureSupported: boolean = Office.context.requirements.isSetSupported('ExcelApi', "1.2");
        var oRange: Excel.Range;

        if (!bFeatureSupported) {
            try {
                oRange = sheet.getCell(1, 1);
                oRange.load('values');
                await ctx.sync();

                var Values = oRange.values;
                oRange.values = Values
                await ctx.sync();
                return true;
            } catch (err) {
                throw 'The sheet is protected.\n Please unprotect the sheet by clicking the "Unprotect sheet" command on Excel "Review" ribbon tab.\n\n' + err;
            }

        }
        else {
            sheet.load(['protection', 'protection/protected']);
            await ctx.sync();
            if (!sheet.protection.protected) {
                bResult = true;
            }
            else {
                try {
                    sheet.protection.unprotect();
                    await ctx.sync();
                    bResult = true;
                }
                catch (err) {
                    throw 'Error occured on unprotecting the sheet.\n\nIs this sheet protected with a password?\n\nTry to unprotect the sheet manually by clicking the "Unprotect sheet" button on the "Review" ribbon tab.\n\n' + err;
                }
            }
        }

        return bResult;
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
    async function onTabChange(event) {
        await Excel.run(async (context) => {
            try {
                if (bInLoadingEvent) return;
                var sActiveTab = event.target.id;
                Office.context.document.settings.set('activetab', sActiveTab);
                await new Promise((resolve, reject) => {
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
                });
                await context.sync();
            }
            catch { }
        }).catch(errorHandler);
    }
})();
