const XLSX = require('xlsx')

const WB = XLSX.readFile('lab/test.xlsx')

console.log(WB.SheetNames)

const sourceWS = WB.Sheets['7.14']

console.log(sourceWS['C5'])

let arr = []

for (let i = 5; i < 26; i++) {
  arr.push(sourceWS[`C${i}`].v)
}

console.log(arr)