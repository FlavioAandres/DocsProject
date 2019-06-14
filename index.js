let docx = require("docx")
let { Table } = require("docx")
let { WidthType } = require("docx")
let fs = require("fs");
let { exampleList } = require('./src/const')
let utils = require('./src/utils')
let doc = new docx.Document();

let newTextStyled = string => new docx.TextRun(string).font('Arial').size(11 * 2)

//Pre Requisites 

const createTable = (boardList) =>{
    let rows = utils.getTotalRows(boardList) + 2

    //Table creation
    let table = new Table(rows, 4);
        table.setWidth(WidthType.PERCENTAGE, '100%')
             .getRow(0)
             .mergeCells(0, 4)
             .createParagraph()
             .addRun(newTextStyled(boardList.BoardName))

    table.getCell(1, 0).createParagraph().addRun(newTextStyled('PROFESSION'))
    table.getCell(1, 1).createParagraph().addRun(newTextStyled('PROFESSION CODE'))
    table.getCell(1, 2).createParagraph().addRun(newTextStyled('SUBJECT AREA'))
    table.getCell(1, 3).createParagraph().addRun(newTextStyled('SUBJECT AREA CODE'))

    return table


}


const writeState = (stateInformation) => {
    doc.createParagraph().addRun(newTextStyled(stateInformation.name).bold())
    doc.createParagraph('\n')
    stateInformation.boards
        .map(board=>createTable(board))
        .map((table,i) => writeBoardTable(table,stateInformation.boards[i]))
        .map(table=>{
            doc.addTable(table)
            doc.createParagraph('\n')
        })
}


//Write Data 

const writeBoardTable = (table, boardInfo) => {
    let initRow = 2
    try {
        boardInfo.professions.map(profession => {
            table.getCell(initRow, 0).createParagraph().addRun(newTextStyled(profession.name))
            table.getCell(initRow, 1).createParagraph().addRun(newTextStyled(profession.code))
            profession.subjectAreas.forEach((subjectArea, i) => {
                table.getCell(initRow, 2).createParagraph().addRun(newTextStyled(subjectArea.code))
                table.getCell(initRow, 3).createParagraph().addRun(newTextStyled(subjectArea.name))
                initRow++
            })
        });
    } catch (error) {
        throw new Error(error)
    }
    return table
}

writeState(exampleList.state)


utils.saveFile(doc)