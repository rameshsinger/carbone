var workSheetBuilder = {
    getWorksheetXML: function (locStr) {
        let resultXml = "";
        let reStr = locStr.replace("*WorksheetTable:", "");
        let workDataList = reStr.split("*@@@*");
        if (workDataList.length >= 3) {
            let workSheetId = workDataList[0];
            console.log(workDataList[1]);
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
                //LDC 104
                case "W15":
                    resultXml += generateParTitle("Test Observation:");
                    resultXml += generateTableStart("LDCTABLE104");
                    resultXml += '<table:table-column table:style-name="LDCTABLE104.A"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE104.B"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE104.C"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE104.D"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE104.E"/>';
                    // HEADER START
                    resultXml += '<table:table-header-rows>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += generateTableCell(["Test Condition"], null, true, true,2,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += generateTableCell(["Parameters"], 3, true, true,null,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += generateTableCell(["Performance"], null, true, true,2,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += '</table:table-row>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += generateTableCell(["Ripple","Frequency","Components (Hz)"], null, true, true,null,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += generateTableCell(["Amplitude of Ripple","Component","(Vrms)"], null, true, true,null,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += generateTableCell(["Time Duration at Condition","(min)"], null, true, true,null,"LDCTABLE104HEADER",'LDCTABLE104BODYCELL');
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '</table:table-row>';
                    resultXml += '</table:table-header-rows>';
                    // HEADER END
                    workSheetRows = worksheetData.sheetDataList;
                    let groupList=[];
                    let groupObject={};
                    for(let i=0;i<workSheetRows.length;i++){
                        if(workSheetRows[i].tcond){
                            if(Object.keys(groupObject).length !== 0){
                                groupList.push(groupObject);
                            }
                            groupObject={firstRow:workSheetRows[i],list:[]}
                        }else if(groupObject.list){
                            groupObject.list.push(workSheetRows[i]);
                        }
                    }
                    if(Object.keys(groupObject).length !== 0){
                        groupList.push(groupObject);
                    }
                    for(let i=0;i<groupList.length;i++){
                        let firstRow = groupList[i].firstRow;
                        let list = groupList[i].list;
                        resultXml += generateTableRowStart('snTableRow');
                        resultXml += generateTableCell([firstRow.tcond], null, false, false,list.length+1,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                        resultXml += generateTableCell([firstRow.rfc], null, false,false,null,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                        resultXml += generateTableCell([firstRow.arc], null, false,false,null,'LDCTABLE104BODY');
                        resultXml += generateTableCell([firstRow.time], null, false, false,list.length+1,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                        resultXml += generateTableCell([getTestStatusValue(firstRow.testStatus)], null, false,false,null,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                        resultXml += '</table:table-row>';
                        for(let j=0;j<list.length;j++){
                            resultXml += generateTableRowStart('snTableRow');
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += generateTableCell([list[j].rfc], null, false,false,null,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                            resultXml += generateTableCell([list[j].arc], null, false,false,null,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += generateTableCell([getTestStatusValue(list[j].testStatus)], null, false,false,null,'LDCTABLE104BODY','LDCTABLE104BODYCELL');
                            resultXml += '</table:table-row>';
                        }
                    }
                    resultXml += '</table:table>';
                    resultXml += generateTableBottomCaption("Test Observation for LDC104", wrkShtTblCnt);
                    break;
                 //LDC 105
                 case "W20":
                    resultXml += generateParTitle("Test Observation:");
                    resultXml += generateTableStart("LDCTABLE105");
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.A"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.B"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.C"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.D"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.E"/>';
                    resultXml += '<table:table-column table:style-name="LDCTABLE105.F"/>';
                    // HEADER START
                    resultXml += '<table:table-header-rows>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += generateTableCell(["Test Condition"], null, true, true,2,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += generateTableCell(["Parameters"], 4, true, true,null,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += generateTableCell(["Performance","Pass/Fail"], null, true, true,2,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += '</table:table-row>';
                    resultXml += generateTableRowStart('snTableRow');
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += generateTableCell(["Steady State","Voltage","(Vdc)"], null, true, true,null,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += generateTableCell(["Voltage Transient","(Vdc)"], null, true, true,null,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += generateTableCell(["Time at Voltage","Transient Level","(msec)"], null, true, true,null,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += generateTableCell(["Oscilloscope Trace","(Vdc vs. Time)"], null, true, true,null,"LDCTABLE105.HEADERPARA",'LDCTABLE105.HEADERCELL');
                    resultXml += '<table:covered-table-cell/>';
                    resultXml += '</table:table-row>';
                    resultXml += '</table:table-header-rows>';
                    // HEADER END
                    // BODY START
                    workSheetRows = worksheetData.sheetDataList;
                    for (let i = 0; i < workSheetRows.length; i++) {
                        if(workSheetRows[i].tcondType){
                            resultXml += generateTableRowStart('snTableRow');
                            resultXml += generateTableCell([workSheetRows[i].tcondType], 6, true, true,null,"LDCTABLE105.SUBHEADERPARA",'LDCTABLE105.SUBHEADERCELL');
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += '<table:covered-table-cell/>';
                            resultXml += '</table:table-row>';
                        }
                        resultXml += generateTableRowStart('snTableRow');
                        resultXml += generateTableCell([workSheetRows[i].tcond], null, false,false,null,'LDCTABLE105.BODYPARA','LDCTABLE105.BODYCELL');
                        resultXml += generateTableCell([workSheetRows[i].svdc], null, false,false,null,'LDCTABLE105.BODYPARA','LDCTABLE105.BODYCELL');
                        resultXml += generateTableCell([workSheetRows[i].voltTr], null, false,false,null,'LDCTABLE105.BODYPARA','LDCTABLE105.BODYCELL');
                        resultXml += generateTableCell([workSheetRows[i].time], null, false,false,null,'LDCTABLE105.BODYPARA','LDCTABLE105.BODYCELL');
                        resultXml += '<table:table-cell table:style-name="LDCTABLE105.BODYCELL" office:value-type="string">';
                        resultXml += drawImage('LDCTABLE105.DRAWPARA','LDCTABLE105.DRAWSTYLE','imageldc105-'+i,'base64','5.62cm','3.999cm')
                        resultXml += '</table:table-cell>';
                        resultXml += generateTableCell([getTestStatusValue(workSheetRows[i].testStatus)], null, false,false,null,'LDCTABLE105.BODYPARA','LDCTABLE105.BODYCELL');
                        resultXml += '</table:table-row>';
                    }
                    // BODY END
                    resultXml += '</table:table>';
                    resultXml += generateTableBottomCaption("Test Observation for LDC105", wrkShtTblCnt);
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
 * @param {Number} rowSpanned (row spanned value)
 * @param {String} paraStyle (paragraph style)
 * @param {String} cellStyle (table cell style)
 */
function generateTableCell(textValueList, colSpanned, isHeader, backgroundColorApply,rowSpanned,paraStyle,cellStyle) {
    let html = '<table:table-cell table:style-name="' + (cellStyle?cellStyle:(backgroundColorApply ? "snTableHeaderCellStyle" : "snTableTableRowCell")) + '" ' + (colSpanned ? (' table:number-columns-spanned="' + colSpanned + '"') : '')+(rowSpanned ? (' table:number-rows-spanned="' + rowSpanned + '"') : '') + '  office:value-type="string">';
    for (let i = 0; i < textValueList.length; i++) {
        html += '<text:p text:style-name="' + (paraStyle ? paraStyle:(isHeader ? 'snTableCellHeader' : 'snTableCellBody')) + '">' + textValueList[i] + '</text:p>';
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

/**
 * @Return Test status string(PASS/FAIL)
 * @param {String} paraStyle (paragraph style)
 * @param {String} drawStyle (svg draw style)
 * @param {String} drawName (svg draw name attribute value)
 * @param {String} base64 (base 64 value)
 * @param {String} width (image width value Eg 30.0cm,30.1in)
 * @param {String} height (image height value Eg 30.0cm,30.1in)
 */
function drawImage(paraStyle,drawStyle,drawName,base64,width,height) {
    let xml = "";
    xml += '<text:p text:style-name="'+paraStyle+'">'
    xml += '<draw:frame draw:style-name="'+drawStyle+'" draw:name="'+drawName+'" svg:width="'+width+'" svg:height="'+height+'"><draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"> <office:binary-data>'
    xml += base64
    xml += "</office:binary-data></draw:image></draw:frame>"
    xml += '</text:p>'
}