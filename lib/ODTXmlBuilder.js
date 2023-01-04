var odtXMLBuilder = {
    imageTextSequence: function (dataStr) {
        let xml = "";
        let dataList = dataStr.split("@@@");
        let prefix = dataList[0];
        let imgCnt = dataList[1];
        let title = dataList[2];
        let imgCaption = dataList[3];
        xml += ('<text:span text:style-name="snImageCaptionSpan">' + prefix + '</text:span>');
        xml += '<text:span text:style-name="snImageCaptionSpan">';
        xml += ('<text:sequence text:ref-name="ref' + imgCaption + imgCnt + '" text:name="' + imgCaption + '" text:formula="ooow:' + imgCaption + '+1" style:num-format="1">' + imgCnt + '');
        xml += '</text:sequence></text:span>';
        xml += '<text:span text:style-name="snImageCaptionSpan">';
        xml += ('<text:s />:' + title);
        xml += '</text:span>';
        return xml;
    },
    emiSuccTableGen: function (dataStr) {
        let xml = "";
        if (dataStr) {
            let dataList = dataStr.split("@@@");
            if (dataList && dataList.length > 0) {
                let jsonStr = dataList[0];
                let titleHeader = dataList[1];
                let tableData = JSON.parse(jsonStr);
                xml += '<table:table  table:style-name="snEmiSucTable">';
                xml += '<table:table-column table:style-name="snEmiSucTableSlColumn"/>';
                xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
                xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
                xml += '<table:table-column table:style-name="snEmiSucTableStatusColumn"/>';
                xml += '<table:table-header-rows>';
                xml += '<table:table-row table:style-name="snEmiSucTableRow">';
                xml += '<table:table-cell table:style-name="snEmiSucTableCell" table:number-columns-spanned="4" office:value-type="string">';
                xml += ('<text:p text:style-name="snEmiSucTableCellHeaderPara">' + titleHeader + '</text:p>');
                xml += '</table:table-cell>';
                xml += '<table:covered-table-cell/>';
                xml += '<table:covered-table-cell/>';
                xml += '<table:covered-table-cell/>';
                xml += '</table:table-row>';
                xml += '<table:table-row table:style-name="snEmiSucTableRow">';
                xml += '<table:table-cell table:style-name="snEmiSucTableCell" office:value-type="string">';
                xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">Sl.</text:p>';
                xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">No</text:p>';
                xml += '</table:table-cell>';
                xml += '<table:table-cell table:style-name="snEmiSucTableCell" office:value-type="string">';
                xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">Name of the Tests</text:p>';
                xml += '</table:table-cell>';
                xml += '<table:table-cell table:style-name="snEmiSucTableCell" office:value-type="string">';
                xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">Limits </text:p>';
                xml += '</table:table-cell>';
                xml += '<table:table-cell table:style-name="snEmiSucTableCell" office:value-type="string">';
                xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">Results</text:p>';
                xml += '</table:table-cell>';
                xml += '</table:table-row>';
                xml += '</table:table-header-rows>';
                for (let i = 0; i < tableData.length; i++) {
                    xml += '<table:table-row table:style-name="snEmiSucTableRowBody">';
                    //Sl No
                    xml += '<table:table-cell table:style-name="snEmiSucTableCellBody" office:value-type="string">';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">' + tableData[i].slNo + '</text:p>';
                    xml += '</table:table-cell>';
                    //Test Name
                    xml += '<table:table-cell table:style-name="snEmiSucTableCellBody" office:value-type="string">';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">' + tableData[i].testName + '</text:p>';
                    xml += '</table:table-cell>';
                    //Limits
                    xml += '<table:table-cell table:style-name="snEmiSucTableCellBody" office:value-type="string">';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">' + tableData[i].limit + '</text:p>';
                    xml += '</table:table-cell>';
                    //Status & Annesxure
                    xml += '<table:table-cell table:style-name="snEmiSucTableCellBody" office:value-type="string">';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">' + tableData[i].testStatus + '</text:p>';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">Refer </text:p>';
                    xml += '<text:p text:style-name="snEmiSucTableCellBodyPara">';
                    xml +=  getAnnexLink(tableData[i].annexLink);
                    xml +=  '</text:p>';
                    xml += '</table:table-cell>';

                    xml += '</table:table-row>';
                }
                xml += '</table:table>';
            }
        }
        return xml;
    },
    uncertainTable: function (dataStr) {
        let xml = "";
        if (dataStr) {
            let tableData = JSON.parse(dataStr);
            xml +=('<text:p text:style-name="uncertainDesc">The following measurement uncertainties are applicable to the relevant tests that are mentioned below:</text:p>');
            xml += '<table:table table:name="TableCustom" table:style-name="TableCustom">';
            xml += '<table:table-column table:style-name="TableCustom.column"/>'
            xml += '<table:table-column table:style-name="TableCustom.column"/>'
            xml += '<table:table-row table:style-name="TableCustom.HeaderRow">'
            xml += '<table:table-cell table:style-name="TableCustom.Header" office:value-type="string"><text:p text:style-name="TableCustom.HeaderP"><text:span text:style-name="TableCustom.HeaderT">Test</text:span></text:p></table:table-cell>'
            xml += '<table:table-cell table:style-name="TableCustom.Header" office:value-type="string"><text:p text:style-name="TableCustom.HeaderP"><text:span text:style-name="TableCustom.HeaderT">Uncertainty (±)</text:span></text:p></table:table-cell>'
            xml += '</table:table-row>'
            for(let i=0;i<tableData.length;i++){
                xml += '<table:table-row table:style-name="TableCustom.cellRow">'
                xml += '<table:table-cell table:style-name="TableCustom.cell" office:value-type="string"><text:p text:style-name="TableCustom.cellP"><text:span text:style-name="TableCustom.cellT">' + tableData[i].testName + '</text:span></text:p></table:table-cell>'
                xml += '<table:table-cell table:style-name="TableCustom.cell" office:value-type="string"><text:p text:style-name="TableCustom.cellP"><text:span text:style-name="TableCustom.cellT">' + tableData[i].uncertain + '</text:span></text:p></table:table-cell>'
                xml += '</table:table-row>'
            }
            xml += '</table:table>'
        }else{
            xml +=('<text:p text:style-name="uncertainDescNonePara">None</text:p>');
        }
        return xml;
    }
}

module.exports = odtXMLBuilder;

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/
function getAnnexLink(jsonStr){
 let xml = "";
 jsonStr = jsonStr.replace("*annexLink:", "");
 let linkData = JSON.parse(jsonStr);
 xml += ('<text:a xlink:type="simple" xlink:href="#'+linkData.id+'" text:style-name="Annex_Internet_20_link" text:visited-style-name="Annex_Visited_20_Internet_20_Link"><text:span text:style-name="Annex_Internet_20_link"><text:span text:style-name="Annex_T36">'+linkData.title+'</text:span></text:span></text:a>');
 return xml;
}