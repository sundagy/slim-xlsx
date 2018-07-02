
const co = require('co');
const slim = require('./index');

co(function *(){

    const xlsx = yield slim.readFile('torg12.xlsx');
    xlsx.shift(1, 35, 10);
    xlsx.copy(1, 34, 35, 10);
    for (let i = 0; i < 11; i++) {
        xlsx.cell(1, 'E' + (34 + i), Math.random());
        xlsx.cell(1, 'B' + (34 + i), i + 1);
    }
    xlsx.cell(1, 'B7', 'ИП "Рога и копыта"');
    yield xlsx.writeFile('torg12-ready.xlsx');

}).catch(err => console.error(err));