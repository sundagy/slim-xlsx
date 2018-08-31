# slim-xlsx

Features:
* Read template xlsx, insert, edit & save (perfect for making invoice-like documents)
* No any library used for xlsx
* Styles & formats are correctly preserving
* Insert new rows with merged cells
* Rows cloning

Install:
```bash
npm i -s https://github.com/sundagy/slim-xlsx.git
```

## Example
```javascript
const slim = require('slim-xlsx');

(async () => {

    // Read document from file
    const xlsx = await slim.readFile('torg12.xlsx');
    // For sheet 1 move rows below 35 down for 10 rows (ow. insert 10 rows in sheet 1 after 35 row)
    xlsx.shift(1, 35, 10);
    // For sheet 1 copy row 34 and insert it starting at row 35 and below 10 times
    xlsx.copy(1, 34, 35, 10);
    // For sheet 1 set value at B & E columns
    for (let i = 0; i < 11; i++) {
        xlsx.cell(1, 'E' + (34 + i), Math.random());
        xlsx.cell(1, 'B' + (34 + i), i + 1);
    }
    // Save to file
    await xlsx.writeFile('torg12-ready.xlsx');
    // return nodebuffer:
    //const buff = await xlsx.getBuffer();

})().catch(console.error);
```
