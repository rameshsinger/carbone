var workSheetBuilder = {
    getWorksheetXML: function (locStr) {
        let resultXml = "";
        let reStr = locStr.replace("*WorksheetTable:", "");
        let workDataList = reStr.split("*@@@*");
        if (workDataList.length >= 3) {
            let workSheetId = workDataList[0];
            let worksheetData = JSON.parse(workDataList[1]);
            let wrkShtTblCnt = workDataList[2];
            let workSheetRows;
            switch (workSheetId) {
                // LDC102
                case "W13":
                    resultXml += generateParTitle("Test Observation:");
                    resultXml += generateTableStart("snBorderTable");
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += '<table:table-column table:style-name="snTableColumn" table:number-columns-repeated="3" />';
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += generateTableCell(["Test Parameter"], 5, true);
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '</table:table-row>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += generateTableCell(["Test Condition"], null, true, true);
                    resultXml += generateTableCell(["Voltage ", "(V dc)"], null, true, true);
                    resultXml += generateTableCell(["Time Duration at condition (min)"], null, true, true);
                    resultXml += generateTableCell(["Re-Start", "(Yes/No)"], null, true, true);
                    resultXml += generateTableCell(["Performance", "Pass/Fail"], null, true, true);
                    resultXml += '</table:table-row>';
                    workSheetRows = worksheetData.sheetDataList;
                    for (let i = 0; i < workSheetRows.length; i++) {
                        resultXml += generateTableRowStart('snTableRow');
                        resultXml += generateTableCell([workSheetRows[i].tcond], null, false);
                        resultXml += generateTableCell([workSheetRows[i].sf], null, false);
                        resultXml += generateTableCell([workSheetRows[i].time], null, false);
                        resultXml += generateTableCell([workSheetRows[i].reStart], null, false);
                        resultXml += generateTableCell([getTestStatusValue(workSheetRows[i].testStatus)], null, false);
                        resultXml += '</table:table-row>';
                    }
                    resultXml += '</table:table>';
                    resultXml += generateTableBottomCaption("Test Observation for LDC102", wrkShtTblCnt);
                    break;
                // LDC 103
                case "W14":
                    resultXml += generateParTitle("Test Observation:");
                    resultXml += generateTableStart("snBorderTable");
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += '<table:table-column table:style-name="snTableColumn"/>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += generateTableCell(["Test Condition"], null, true, true);
                    resultXml += generateTableCell(["Frequency of Voltage Distortion (Hz)"], null, true, true);
                    resultXml += generateTableCell(["Amplitude of Voltage Distortion (Vrms)"], null, true, true);
                    resultXml += generateTableCell(["Time Duration at Condition<text:line-break/>(min)"], null, true, true);
                    resultXml += generateTableCell(["Performance"], null, true, true);
                    resultXml += '</table:table-row>';
                    workSheetRows = worksheetData.sheetDataList;
                    for (let i = 0; i < workSheetRows.length; i++) {
                        resultXml += generateTableRowStart('snTableRow');
                        resultXml += generateTableCell([workSheetRows[i].tcond], null, false);
                        resultXml += generateTableCell([workSheetRows[i].sa], null, false);
                        resultXml += generateTableCell([workSheetRows[i].ab], null, false);
                        resultXml += generateTableCell([workSheetRows[i].time], null, false);
                        resultXml += generateTableCell([getTestStatusValue(workSheetRows[i].testStatus)], null, false);
                        resultXml += '</table:table-row>';
                    }
                    resultXml += '</table:table>';
                    resultXml += generateTableBottomCaption("Test Observation for LDC103", wrkShtTblCnt);
                    break;
            }
        }
        return resultXml;
    }
}

module.exports = workSheetBuilder;

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/

/**
 * Generate Paragrah Title with Bold and underscore
 * @param {String} title (Title Text)
 */
function generateParTitle(title) {
    return '<text:p text:style-name="snParagraphTitle">' + title + '</text:p>';
}

/**
 * Generate Table start
 * Available styles ::: snBorderTable
 * @param {String} cssStyle (Style Name)
 */
function generateTableStart(cssStyle) {
    return '<table:table table:name="TableCustom" table:style-name="' + cssStyle + '">';
}

/**
 * Generate Table Row start
 * Available stytles ::: snTableRow
 * @param {String} cssStyle (Style Name)
 */
function generateTableRowStart(cssStyle) {
    return '<table:table-row table:style-name="' + cssStyle + '">';
}

/**
 * Generate Table Cell
 * Available stytles ::: snTableRow
 * @param {Array<String>} textValueList (Cell Value List)
 * @param {Number} colSpanned (Col spanned value)
 * @param {boolean} isHeader (Header cell == true, body cell == false)
 * @param {boolean} backgroundColorApply (Apply back ground color brown == true or non null, to not apply == null or false or "")
 */
function generateTableCell(textValueList, colSpanned, isHeader, backgroundColorApply) {
    let html = '<table:table-cell table:style-name="' + (backgroundColorApply ? "snTableHeaderCellStyle" : "snTableTableRowCell") + '" ' + (colSpanned ? (' table:number-columns-spanned="' + colSpanned + '"') : '') + '  office:value-type="string">';
    for (let i = 0; i < textValueList.length; i++) {
        html += '<text:p text:style-name="' + (isHeader ? 'snTableCellHeader' : 'snTableCellBody') + '">' + textValueList[i] + '</text:p>';
    }
    html += '</table:table-cell>';
    return html;
}

/**
 * Generate Table Bottom Caption
 * Available stytles ::: snTableRow
 * @param {String} titleText (Caption name)
 * @param {String} tblCnt (Table Number)
 */
function generateTableBottomCaption(titleText, tblCnt) {
    return '<text:p text:style-name="snTableCaptionPara"/><text:p text:style-name="snTableCaptionPara"><text:span text:style-name="snTableCaptionText">Table ' + tblCnt + ': ' + titleText + '</text:span></text:p>';
}

/**
 * @Return Test status string(PASS/FAIL)
 * @param {String} testStatusCode (Possible values ::: 1,2,3)
 */
function getTestStatusValue(testStatusCode) {
    if (testStatusCode == 1) {
        return "PASS"
    }
    else if (testStatusCode == 2) {
        return "FAIL"
    }
    else if (testStatusCode == 3) {
        return "NA"
    }
    else {
        return ""
    }
}