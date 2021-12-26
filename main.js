const fs = require('fs');
const os = require('os');
const path = require('path');
const zl = require("zip-lib");

const APP_PREFIX = 'tux-docx-template';
const DOCUMENT_FILE = 'word/document.xml'
const START_TEXT='<w:r><w:t>';
const START_TEXT_BOLD='<w:r w:rsidRPr=""><w:rPr><w:b/><w:bCs/></w:rPr><w:t>'
const END_TEXT='</w:t></w:r>';
const REGEX_RSIDRDefault = /w:rsidRDefault="([^"]+)"/g;

class DocxTemplate {

    constructor(filename) {
        this.filename=filename;
        this.tmpDir=null;
        this.document=null;
        this.rsidDefaults = [];
    }

    replaceInDocument(from,toData) {
        let idx=null;
        toData=toData.replace(/(?:(<)(?!(b|\/b)))/ig,'<w:r><w:t>&gt;</w:t></w:r>$2').replace(/(?:([^b])(>))/ig,'$1<w:r><w:t>&lt;</w:t></w:r>');
        if (toData.indexOf('<b>')<0) {
            this.document=this.document.replaceAll(from,toData);
            return;
        }
        while ((idx = this.document.indexOf(from))>=0) {
            let regularCharactersBefore=0;
            let regularCharactersAfter=0;
            while (this.document[idx-1-regularCharactersBefore] !== '>') {
                regularCharactersBefore++;
            }
            while (this.document[idx+1+regularCharactersAfter] !== '<') {
                regularCharactersAfter++;
            }
            const allString = this.document.slice(idx-regularCharactersBefore,idx+regularCharactersAfter+1);
            let modifiedString = '';
            let leftoverString = allString;
            while (leftoverString.length>0) {
                let StartTagIdx = leftoverString.indexOf('{');
                if (StartTagIdx<0) {
                    modifiedString+=START_TEXT+leftoverString+END_TEXT;
                    leftoverString='';
                } else if (StartTagIdx === 0) {
                    const endTagIdx = leftoverString.indexOf('}');
                    modifiedString+=leftoverString.slice(0,endTagIdx+1);
                    leftoverString=leftoverString.slice(endTagIdx+1);
                } else {
                    const regularTxt = leftoverString.slice(0,StartTagIdx);
                    modifiedString+=START_TEXT + regularTxt + END_TEXT;
                    leftoverString=leftoverString.slice(StartTagIdx);
                }
            }
            const sid = this.getRSidForPos(idx);
            let to = toData;
            let startBTagFirst = null;
            let toParsed = '';
            while ((startBTagFirst=to.indexOf('<b>'))>0) {
                const endBTagPos=to.indexOf('</b>');
                const beforeB = to.slice(0,startBTagFirst);
                const afterB = to.slice(startBTagFirst+3,endBTagPos);
                toParsed+=START_TEXT+beforeB+END_TEXT+START_TEXT_BOLD.replace('""','"' + sid + '"') +afterB+END_TEXT;
                to=to.slice(endBTagPos+4);
            }
            if (to.length>0) {
                toParsed+=START_TEXT+to+END_TEXT;
            }

            const resultString = modifiedString.replaceAll(from,toParsed);
            this.document=this.document.replace(allString,resultString);
        }
    }

    replaceTemplateObject(obj) {
        Object.keys(obj).forEach((title)=>{
            const key = `{${title}}`;
            const val = obj[title];
            this.replaceInDocument(key,val);
        });
    }

    async zipDocument(filename) {
        await zl.archiveFolder(this.tmpDir,filename);
    }

    getRSIDRDefaults() {
        const matches = Array.from(this.document.matchAll(REGEX_RSIDRDefault));
        for (const match of matches) {
            const sid = match[1];
            const pos = match.index;
            this.rsidDefaults.push({sid,pos});
        }
    }

    getRSidForPos(pos) {
        let ret = null;
        for (let i=0;i<this.rsidDefaults.length;i++) {
            if (i === 0) {
                ret=this.rsidDefaults[0].sid;
            } else {
                if (this.rsidDefaults[i].pos<pos) {
                    ret=this.rsidDefaults[i].sid;
                }
            }
        }
        return ret;
    }

    saveDocument() {
        fs.writeFileSync(this.tmpDir + '/' + DOCUMENT_FILE,this.document);
    }

    readDocument() {
        this.document=fs.readFileSync(this.tmpDir + '/' + DOCUMENT_FILE).toString();
    }

    async extractFile() {
        await zl.extract(this.filename,this.tmpDir);
        console.debug(`extracted ${this.filename} into ${this.tmpDir}`);
    }

    createTmpDir() {
        this.tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), APP_PREFIX));
        console.debug(`created tmp dir ${this.tmpDir}`);
    }

    deleteTmpDir() {
        fs.rmSync(this.tmpDir, { recursive: true });
    }
myt
    async renderAndSave(obj, saveFile) {
        this.createTmpDir();
        await this.extractFile();
        this.readDocument();
        this.getRSIDRDefaults();
        this.replaceTemplateObject(obj);
        this.saveDocument();
        await this.zipDocument(saveFile);
        this.deleteTmpDir();
    }
}

async function main() {
    const file = new DocxTemplate('mytest.docx');
    const obj = {'FOO':'Hello<b>How are ya</b>'};
    await file.renderAndSave(obj,'mytest_result.docx');
    console.info('done');
}

main().then(()=>{
    console.log('done');
}).catch((e)=>{
    console.log(e.message);
})
