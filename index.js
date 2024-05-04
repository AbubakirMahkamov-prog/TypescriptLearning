const readXlsxFile = require('read-excel-file/node');
const writeXlsxFile = require('write-excel-file/node')

const schema = [
    {
      column: 'Name',
      type: String,
      value: val => val.name
    },
    {
        column: 'Shots',
        type: String,
        value: val => val.shot
    },
    {
        column: 'All summ',
        type: String,
        value: val => val.all_sum
    },
]
  
const DATA = [];

let employees = [];
(async () => {
    const rows = await readXlsxFile('./qxodim.xlsx');
    const moneys = await readXlsxFile('./qmoney.xlsx');
    for (const item of rows) {
        const inx = employees.findIndex((val) => val.name == item[1]);
        if (inx === -1) {
            employees.push({
                name: item[1],
                shots: [item[0]]
            })
        } else {
            employees[inx].shots.push(item[0])
        }
    }
    for (const employee of employees) {
        let allSumm = 0;
        for (const shot of employee.shots) {
            const ownMoneys = moneys.filter((temp) => {
                if (shot == temp[0]) {
                    return true;
                }
                return false;
            })
            let summ = 0;
            for (const money of ownMoneys) {
                summ += Number(money[1])
            }
            allSumm += summ;
        }
        DATA.push({
            name: employee.name,
            shot: JSON.stringify(employee.shots),
            all_sum: JSON.stringify(allSumm)
        })
    }
    await writeXlsxFile(DATA, {
        schema,
        filePath: 'qresult.xlsx'
      })
    // console.log(employees)
})()
