const fs = require('fs');

const _ = require('underscore');
const xml2js = require('xml2js');
const JSZip = require('jszip');
const util = require('util');

const fs_readFile = util.promisify(fs.readFile);
const fs_writeFile = util.promisify(fs.writeFile);
const parseString = util.promisify(xml2js.parseString);

const fcopy = (json) => {
    return JSON.parse(JSON.stringify(json));
};

const shift = function (sheetId, row, count) {
    const xlsx = this;
    let sheet = xlsx.xl.worksheets['sheet' + sheetId];

    row = parseInt(row);
    const rows = sheet.worksheet.sheetData[0].row.filter(a => parseInt(a.$.r) >= row);
    rows.sort((a, b) => a.$.r - b.$.r);

    for (let r of rows) {
        r.$.r = '' + (parseInt(r.$.r) + count);
        if (!r.c) continue;
        for (let c of r.c) {
            let rn = parseInt(c.$.r.match(/\d+/)[0]);
            rn += count;
            c.$.r = c.$.r.replace(/\d+/, rn);
        }
    }
    for (let mc of sheet.worksheet.mergeCells[0].mergeCell) {
        let from = mc.$.ref.split(':')[0];
        let to = mc.$.ref.split(':')[1];

        let from_r = parseInt(from.match(/\d+/)[0]);
        let to_r = parseInt(to.match(/\d+/)[0]);

        if (from_r >= row) {
            from_r += count;
            from = from.replace(/\d+/, from_r);
        }
        if (to_r >= row) {
            to_r += count;
            to = to.replace(/\d+/, to_r);
        }

        mc.$.ref = from + ':' + to;
    }
    for (let c of xlsx.xl.calcChain.calcChain.c) {
        if (c.$.i != sheetId) continue;
        let r = parseInt(c.$.r.match(/\d+/)[0]);
        if (r >= row) {
            r += count;
            c.$.r = c.$.r.replace(/\d+/, r);
        }
    }
};

const copy = function (sheetId, row, set, count) {
    const xlsx = this;
    let sheet = xlsx.xl.worksheets['sheet' + sheetId];

    let sourceRowId = sheet.worksheet.sheetData[0].row.findIndex(a => parseInt(a.$.r) === row);
    const sourceRow = sheet.worksheet.sheetData[0].row[sourceRowId];
    sourceRowId += 1;
    const sourceMerge = [];
    for(let mc of sheet.worksheet.mergeCells[0].mergeCell){
        let from = mc.$.ref.split(':')[0];
        let to = mc.$.ref.split(':')[1];

        let from_r = parseInt(from.match(/\d+/)[0]);
        let to_r = parseInt(to.match(/\d+/)[0]);

        if (from_r <= row && to_r >= row) {
            const nmc = fcopy(mc);
            sourceMerge.push(nmc);
        }
    }
    for (let r = set; r < set + count; ++r) {
        sheet.worksheet.sheetData[0].row.filter(a => parseInt(a.$.r) !== r);
        let nr = fcopy(sourceRow);
        nr.$.r = r;
        for (let c of nr.c) {
            c.$.r = c.$.r.replace(row, r);
        }
        sheet.worksheet.sheetData[0].row.splice(sourceRowId, 0, nr);
        sourceRowId += 1;

        for (let mc of fcopy(sourceMerge)) {
            const s = mc.$.ref.split(':');
            mc.$.ref = s[0].replace(/\d+/, r) + ':' + s[1].replace(/\d+/, r);
            sheet.worksheet.mergeCells[0].mergeCell.push(mc);
        }
    }

};

const getBuffer = function *(fn){
    const zip = this._.zip;
    for (let b of this._.back) {
        const builder = new xml2js.Builder();
        const destXml = builder.buildObject(b[1]);
        zip.file(b[0], destXml);
    }
    return yield zip.generateAsync({type: "nodebuffer"});
    yield fs_writeFile(fn, destData);
};

const writeFile = function *(fn){
    const destData = yield this.getBuffer();
    yield fs_writeFile(fn, destData);
};

const abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

const chid = (addr) => {
    const ch = addr.match(/[A-Z]+/)[0].split('');
    let result = 0;
    for (let i in ch) {
        const c = ch[i];
        result += (abc.indexOf(c) + 1) + (25 * (ch.length - i - 1));
    }
    return result;
};

const cell = function(sheetId, addr, value){
    let sheet = this.xl.worksheets['sheet' + sheetId];
    const r = parseInt(addr.match(/\d+/)[0]);
    let row = sheet.worksheet.sheetData[0].row.find(a => parseInt(a.$.r) === r);
    if (!row) {
        row = {
            '$': {
                r: '' + r,
                spans: '2:58',
                customFormat: '1',
                ht: '21',
                customHeight: '1',
                thickBot: '1',
                'x14ac:dyDescent': '0.25' },
            c: [ { '$': {r: addr}, v: [''] } ]
        };
        sheet.worksheet.sheetData[0].row.push(row);
        sheet.worksheet.sheetData[0].row.sort((a, b) => a.$.r - b.$.r);
    }
    if (!row.c) {
        row.c = [ { '$': {r: addr}, v: [''] }];
    }
    let cell = row.c.find(a => a.$.r === addr);
    if (cell) {
        cell.v = ['' + value];
    } else {
        row.c.push({ '$': {r: addr}, v: ['' + value] });
        row.c.sort((a, b) => chid(a.$.r) - chid(b.$.r));
    }
};

exports.readFile = function*(fn){
    const data = yield fs_readFile(fn);
    const zip = yield JSZip.loadAsync(data);
    let xlsx = {_:{zip},
        writeFile,
        shift,
        copy,
        cell,
    };
    let back = [];
    for (let f of _.keys(zip.files)) {
        if (f.match(/\.xml$/)) {
            const p = f.split('/');
            let last = xlsx;
            for(let pp of p.slice(0, p.length - 1)) {
                last[pp] = last[pp] || {};
                last = last[pp];
            }
            const xml = yield zip.file(f).async("string");
            const json = yield parseString(xml);
            back.push([f, json]);
            last[_.last(p).replace('.xml','')] = json;
        }
    }
    xlsx._.back = back;
    return xlsx;
};