let docx = require("docx")
let { Table, TableCellBorders } = require("docx")
let { WidthType } = require("docx")
let fs = require("fs");
let { exampleList, textConstants } = require('./src/const')
let utils = require('./src/utils')
let doc = new docx.Document();

const textStyled = (string, size = 11, color = '000') => new docx.TextRun(string).font('Arial').size(size * 2).color(color)
const emptyBorders = () => new TableCellBorders().addBottomBorder('none', 0, 'white').addTopBorder('none', 0, 'white')
const fullBorders = () => new TableCellBorders().addBottomBorder().addEndBorder().addTopBorder('none', 0, 'white')

//Pre Requisites 
const createHeadersAndFooters = () => {
    doc.Header.createParagraph(
        textStyled(textConstants.header, 14, '6b0b20')
        .bold()
        .underline("single", "6b0b20")
    );
    doc.Footer.createParagraph(
        textStyled(textConstants.footer, 10, '6b0b20')
        .bold().pageNumber()
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
        .addRun(textStyled(boardList.BoardName).bold())

    table.getCell(1, 0).createParagraph()
        .addRun(textStyled('Profession').bold())

    table.getCell(1, 1).createParagraph()
        .addRun(textStyled('Profession Code').bold())

    table.getCell(1, 2).createParagraph()
        .addRun(textStyled('Subject Area').bold())

    table.getCell(1, 3).createParagraph()
        .addRun(textStyled('Subject Area code').bold())

    return table
}

//Write Data 

const writeBoardTable = (table, boardInfo) => {
    let initRow = 2
    try {
        boardInfo.professions.map(profession => {
            let pointer = 2
            table.getCell(initRow, 0).Properties.setWidth('40%', WidthType.PERCENTAGE)
            table.getCell(initRow, 0).createParagraph().addRun(textStyled(profession.name))
            table.getCell(initRow, 1).createParagraph().addRun(textStyled(profession.code))
            profession.subjectAreas.forEach(subjectArea => {
                table.getCell(initRow, 0).root[0].root[0] = emptyBorders()
                table.getCell(initRow, 1).root[0].root[0] = emptyBorders()
                if (profession.subjectAreas.length + 1 === pointer) {
                    table.getCell(initRow, 1).root[0].root[0] = fullBorders()
                    table.getCell(initRow, 0).root[0].root[0] = fullBorders()
                }
                table.getCell(initRow, 2).createParagraph().addRun(textStyled(subjectArea.code))
                table.getCell(initRow, 3).createParagraph().addRun(textStyled(subjectArea.name))
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

//main function 
const writeState = (stateInformation) => {
    createHeadersAndFooters()
    doc.createParagraph().addRun(textStyled(stateInformation.name, 14).bold())
    doc.createParagraph('\n')
    stateInformation.boards
        .map(board => createTable(board))
        .map((table, i) => writeBoardTable(table, stateInformation.boards[i]))
        .map(writedTable => {
            doc.addTable(writedTable)
            doc.createParagraph('\n')
        })
}



writeState(exampleList.state)


utils.saveFile(doc)