const unzipper = require("unzipper")
const fs = require('fs');
const { readFile, writeFile } = require('fs/promises')
var zipper = require('zip-local');

var docxXMLEditor = {
    /**
    * @return String
    * @param {String} docxFilePath (word document path(doxx extension file))
    * @param {Number} omitPagesCount (pages count to substract from total pages count)
    */
    updateTotalPageNumber: async function (docxFilePath, omitPagesCount) {
        var filename = docxFilePath.replace(/^.*[\\\/]/, '');
        var extension = filename.split(".")[1];
        var newDirFullPath = docxFilePath.replace("." + extension, '');
        fs.mkdirSync(newDirFullPath);
        zipper.sync.unzip(docxFilePath).save(newDirFullPath);
        var appXmlPath = newDirFullPath + "/docProps/app.xml";
        const appXmlContent = await getFileContent(appXmlPath);
        var totalPageCount = Number(getXmlNodeValue(appXmlContent, 'Pages')) - omitPagesCount;
        var fileNameList = fs.readdirSync(newDirFullPath + "/word");
        for (let i = 0; i < fileNameList.length; i++) {
            let fileNme = fileNameList[i];
            if(fileNme.startsWith('header')){
                await replaceFileContent(newDirFullPath + "/word/" + fileNme, 'xxxPAGExxx', totalPageCount);
                // console.log("change done to:" + fileNme);
            }
        }
        fs.unlinkSync(docxFilePath);
        zipper.sync.zip(newDirFullPath).compress().save(docxFilePath);
        // console.log("Zip done")
        fs.rmSync(newDirFullPath, { recursive: true, force: true });
        return "success";
    }
}

module.exports = docxXMLEditor;

/** ***************************************************************************************************************/
/* Privates methods */
/** ***************************************************************************************************************/

/**
 * Returns the content of the given file `path` as a string.
 * @param {String} path
 * @returns {String}
 */
async function getFileContent(path) {
    return await readFile(path, 'utf8')
}

/**
 * Returns value between xml tag from content string
 * @param {String} content
 * @param {String} xmlTagName xml tag name Eg. body,head
 * @returns {String} value between xml tag
 */
function getXmlNodeValue(content, xmlTagName) {
    var re = new RegExp('<' + xmlTagName + '>([^<]*)<\/' + xmlTagName + '>', '');
    var match = content.match(re);
    return match[1];
}

/**
 * Replace file content and write again
 * @param {String} filePath file path
 * @param {String} replaceSource replace source 
 * @param {String} replaceTarget replace target
 */
async function replaceFileContent(filePath, replaceSource, replaceTarget) {
    const data = await getFileContent(filePath);
    var regex = new RegExp(replaceSource, "g");
    const result = data.replace(regex, replaceTarget);
    await writeFile(filePath, result, 'utf8');
}