let docx = require("docx")
let { Table } = require("docx")
let { WidthType } = require("docx")
let fs = require("fs");
let { exampleList } = require('./src/const')
let utils = require('./src/utils')
let doc = new docx.Document();


//Pre Requisites 
let rows = utils.getTotalRows(exampleList) + 2
    //Table creation
let table = new Table(rows, 4);
table.setWidth(WidthType.PERCENTAGE, '98%')
console.log(`Total Rows: ${rows}`)
let firstRow = table.getRow(0)
let cell = firstRow.mergeCells(0, 4)
cell.createParagraph(exampleList.BoardName)

//constants 
table.getCell(1, 0).createParagraph('PROFESSION')
table.getCell(1, 1).createParagraph('PROFESSION CODE')
table.getCell(1, 2).createParagraph('SUBJECT AREA')
table.getCell(1, 3).createParagraph('SUBJECT AREA CODE')

//Write Data 

const writeBoard = (list) => {
    let initRow = 2
    let subjectCount = 0
    try {
        list.professions.forEach(profession => {
            table.getCell(initRow, 0).createParagraph(profession.name)
            table.getCell(initRow, 1).createParagraph(profession.code)
            let subjectAreas = profession.subjectAreas

            subjectAreas.forEach((subjectArea, i) => {
                    subjectCount++
                    table.getCell(initRow, 2).createParagraph(subjectArea.code)
                    table.getCell(initRow, 3).createParagraph(subjectArea.name)
                    initRow++
                })
                // table.getColumn(0).mergeCells(initRow, initRow + subjectCount - 1)
        });
    } catch (error) {
        console.error('hubo un error: ' + error)
    }
}

writeBoard(exampleList)

doc.addTable(table)

utils.saveFile(doc)