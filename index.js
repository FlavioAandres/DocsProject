let docx = require("docx")
let { Table, TableCellBorders } = require("docx")
let { WidthType } = require("docx")
let fs = require("fs");
let { exampleList, textConstants } = require('./src/const')
let utils = require('./src/utils')
let doc = new docx.Document();

const newTextStyled = (string, size = 11, color = '000') => new docx.TextRun(string).font('Arial').size(size * 2).color(color)
const setEmptyBorder = () => new TableCellBorders()
    .addBottomBorder('none', 0, 'white')
    .addTopBorder('none', 0, 'white')

const setFullBorder = () => new TableCellBorders().addBottomBorder().addEndBorder().addTopBorder('none', 0, 'white')

//Pre Requisites 
const createHeadersAndFooters = () => {
    doc.Header.createParagraph(
        newTextStyled(textConstants.header, 14, '6b0b20')
        .bold()
        .underline("single", "6b0b20")
    );
    doc.Footer.createParagraph(
        newTextStyled(textConstants.footer, 10, '6b0b20')
        .bold()
    );
}

const createTable = (boardList) => {
    let rows = utils.getTotalRows(boardList) + 2

    //Table creation
    let table = new Table(rows, 4);
    table.setWidth(WidthType.PERCENTAGE, '100%')
        .getRow(0)
        .mergeCells(0, 4)
        .createParagraph()
        .addRun(newTextStyled(boardList.BoardName).bold())

    table.getCell(1, 0).createParagraph()
        .addRun(newTextStyled('Profession').bold())

    table.getCell(1, 1).createParagraph()
        .addRun(newTextStyled('Profession Code').bold())

    table.getCell(1, 2).createParagraph()
        .addRun(newTextStyled('Subject Area').bold())

    table.getCell(1, 3).createParagraph()
        .addRun(newTextStyled('Subject Area code').bold())

    return table
}


const writeState = (stateInformation) => {
    createHeadersAndFooters()
    doc.createParagraph().addRun(newTextStyled(stateInformation.name, 14).bold())
    doc.createParagraph('\n')
    stateInformation.boards
        .map(board => createTable(board))
        .map((table, i) => writeBoardTable(table, stateInformation.boards[i]))
        .map(table => {
            doc.addTable(table)
            doc.createParagraph('\n')
        })
}


//Write Data 

const writeBoardTable = (table, boardInfo) => {
    let initRow = 2
    try {
        boardInfo.professions.map(profession => {
            let pointer = 2
            table.getCell(initRow, 0).createParagraph().addRun(newTextStyled(profession.name))
            table.getCell(initRow, 1).createParagraph().addRun(newTextStyled(profession.code))
            profession.subjectAreas.forEach(subjectArea => {
                table.getCell(initRow, 0).root[0].root[0] = setEmptyBorder()
                table.getCell(initRow, 1).root[0].root[0] = setEmptyBorder()
                if (profession.subjectAreas.length + 1 === pointer) {
                    table.getCell(initRow, 1).root[0].root[0] = setFullBorder()
                    table.getCell(initRow, 0).root[0].root[0] = setFullBorder()
                }
                table.getCell(initRow, 2).createParagraph().addRun(newTextStyled(subjectArea.code))
                table.getCell(initRow, 3).createParagraph().addRun(newTextStyled(subjectArea.name))
                initRow++
                pointer++
            })
        });
    } catch (error) {
        console.error(error)
        return table
    }
    return table
}

writeState(exampleList.state)


utils.saveFile(doc)