let docx = require("docx")
let fs = require("fs");


const getTotalRows = list => {
    let rows = 0
    list.professions.forEach(profession => {
        rows += Object.keys(profession.subjectAreas).length
    });
    return rows
}

const saveFile = async(doc) => {
    const packer = new docx.Packer();
    let buffer = await packer.toBuffer(doc)
    fs.writeFileSync("My Document.docx", buffer);
}


module.exports = {
    getTotalRows,
    saveFile,
}