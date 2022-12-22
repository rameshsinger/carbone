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
    }
}

module.exports = odtXMLBuilder;

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/
