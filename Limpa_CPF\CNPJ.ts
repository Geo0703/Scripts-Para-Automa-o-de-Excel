function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();
    let column = range.getColumn(0); // Altere o índice da coluna conforme necessário (0 = coluna A)

    column.getValues().forEach((row, rowIndex) => {
        let cellValue = row[0].toString();
        let cleanedValue = cellValue.replace(/[.\-\/]/g, '');
        column.getCell(rowIndex, 0).setValue(cleanedValue);
    });
}
