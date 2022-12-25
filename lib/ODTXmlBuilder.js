var odtXMLBuilder = {
    imageTextSequence:function(dataStr){
        let xml = "";
        let dataList = dataStr.split("@@@");
        let prefix = dataList[0];
        let imgCnt = dataList[1];
        let title = dataList[2];
        let imgCaption = dataList[3];
        xml += ('<text:span text:style-name="snImageCaptionSpan">'+prefix+'</text:span>');
        xml += '<text:span text:style-name="snImageCaptionSpan">';
        xml += ('<text:sequence text:ref-name="ref'+imgCaption+imgCnt+'" text:name="'+imgCaption+'" text:formula="ooow:'+imgCaption+'+1" style:num-format="1">'+imgCnt+'');
        xml += '</text:sequence></text:span>';
        xml += '<text:span text:style-name="snImageCaptionSpan">';
        xml += ('<text:s />:'+title);
        xml += '</text:span>';
        return xml;
    },
    emiSuccTableGen:function(dataStr){
        let xml = "";
        if(dataStr){
           let dataList = dataStr.split("@@@");
           if(dataList && dataList.length > 0){
            let jsonStr = dataList[0];
            let titleHeader = dataList[1];
            let tableData = JSON.parse(jsonStr);
            xml += '<table:table  table:style-name="snEmiSucTable">';
            xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
            xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
            xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
            xml += '<table:table-column table:style-name="snEmiSucTableColumn"/>';
            xml += '<table:table-header-rows>';
            xml += '<table:table-row table:style-name="snEmiSucTableRow">';
            xml += '<table:table-cell table:style-name="snEmiSucTableCell" table:number-columns-spanned="4" office:value-type="string">';
            xml += ('<text:p text:style-name="snEmiSucTableCellHeaderPara">'+titleHeader+'</text:p>');
            xml += '</table:table-cell>';
            xml += '<table:covered-table-cell/>';
            xml += '<table:covered-table-cell/>';
            xml += '<table:covered-table-cell/>';
            xml += '</table:table-row>';
            xml += '<table:table-row table:style-name="snEmiSucTableRow">';
            xml += '<table:table-cell table:style-name="snEmiSucTableCell" office:value-type="string">';
            xml += '<text:p text:style-name="snEmiSucTableCellHeaderPara">Sl. No</text:p>';
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
           }
        }
        return xml;
    }
}

module.exports = odtXMLBuilder;

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/
