const fs = require('fs');
const os = require('os');
const path = require('path');
const zl = require("zip-lib");

const APP_PREFIX = 'tux-docx-template';
const CHARTS_DIR = 'word/charts';
const DOCUMENT_FILE = 'word/document.xml';
const START_TEXT='<w:r><w:t>';
const START_TEXT_BOLD='</w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">';
const END_TEXT='</w:t></w:r>';
const NEW_LINE='</w:t></w:r><w:r><w:br/><w:t xml:space="preserve">';
const END_TEXT_BOLD='</w:t></w:r><w:r><w:t xml:space="preserve">';
const REGEX_MAYBE_STUFF = '(?:<[^>]+>)*'

class DocxTemplate {

    constructor(filename) {
        this.filename=filename;
        this.tmpDir=null;
        this.document=null;
        this.rsidDefaults = [];
    }

    replaceInDocument(from,toData) {
        let idx=null;
        let fromRegexPattern='';
        for (let i=0;i<from.length;i++) {
            fromRegexPattern+=from[i];
            if (i+1 < from.length) {
                fromRegexPattern+=REGEX_MAYBE_STUFF;
            }
        }
        let fromRegex = new RegExp(fromRegexPattern,'g');
        toData=toData.toString()
            .replace(/(\s*)<b>/g,'<b>$1')
                .replace(/(?:(<)(?!(b|\/b)))/ig,'<w:r><w:t>&gt;</w:t></w:r>$2').replace(/(?:([^b])(>))/ig,'$1<w:r><w:t>&lt;</w:t></w:r>').
        replace(/(\s*)<b>/g, START_TEXT_BOLD).replaceAll('</b>',END_TEXT_BOLD)
            .replaceAll('\n',NEW_LINE);
        this.document=this.document.replaceAll(from,toData);
        this.document=this.document.replaceAll(fromRegex,toData);
    }

    replaceTemplateObject(obj) {
        Object.keys(obj).forEach((title)=>{
            const key = `{${title}}`
            const val = obj[title];
            this.replaceInDocument(key,val);
        });
    }

    async zipDocument(filename) {
        await zl.archiveFolder(this.tmpDir,filename);
    }

    saveDocument() {
        fs.writeFileSync(this.tmpDir + '/' + DOCUMENT_FILE,this.document);
    }

    saveChartFile(data,file) {
        fs.writeFileSync(this.tmpDir + '/' + CHARTS_DIR + '/' + file ,data);
    }

    readDocument() {
        this.document=fs.readFileSync(this.tmpDir + '/' + DOCUMENT_FILE).toString();
    }
    readChartFile(file) {
        return fs.readFileSync(this.tmpDir + '/' + CHARTS_DIR + '/' + file).toString();
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

    async replaceInChartFile(chartNumber, chartProps) {
        const xlsFile = this.tmpDir + '/word/embeddings/Microsoft_Excel_Worksheet' +
            ((chartNumber>1) ? (chartNumber-1) : '') + '.xlsx';
        const xlsTmpDir = fs.mkdtempSync(path.join(os.tmpdir(), APP_PREFIX));
        await zl.extract(xlsFile,xlsTmpDir);
        const sharedStringsFilename = xlsTmpDir + '/xl/sharedStrings.xml';
        const worksheetFilename = xlsTmpDir + '/xl/worksheets/sheet1.xml';
        let sharedStringsData = fs.readFileSync(sharedStringsFilename).toString();
        let worksheetData = fs.readFileSync(worksheetFilename).toString();
        const filename = 'chart' + chartNumber + '.xml';
        let file = this.readChartFile(filename);
        for (let key in chartProps) {
            const prop = '>' + key + '<';
            const value='>' + chartProps[key] + '<';
            file=file.replaceAll(prop,value);
            sharedStringsData=sharedStringsData.replaceAll(prop,value);
            worksheetData=worksheetData.replaceAll(prop,value);
        }
        this.saveChartFile(file,filename);
        fs.writeFileSync(sharedStringsFilename,sharedStringsData);
        fs.writeFileSync(worksheetFilename,worksheetData);
        await zl.archiveFolder(xlsTmpDir,xlsFile);
    }

    async renderAndSave(obj, chartObj,saveFile) {
        this.createTmpDir();
        await this.extractFile();
        if (obj) {
            this.readDocument();
            this.replaceTemplateObject(obj);
            this.saveDocument();
        }
        if (chartObj) {
            for (let chartFile in chartObj) {
                await this.replaceInChartFile(chartFile,chartObj[chartFile]);
            }
        }
        await this.zipDocument(saveFile);
        this.deleteTmpDir();
    }
}

async function renderAndSaveDocx(fromFile,objProps,chartsObj,toFile) {
    const file = new DocxTemplate(fromFile);
    await file.renderAndSave(objProps,chartsObj,toFile);
}

module.exports.renderAndSaveDocx=renderAndSaveDocx;

