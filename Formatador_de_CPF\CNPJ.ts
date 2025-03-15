function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getRange("J:J"); // Altere "A:A" para a coluna desejada
    let values = range.getValues();

    for (let i = 0; i < values.length; i++) {
        let cellValue = values[i][0].toString();
        if (cellValue.length === 11) {
            values[i][0] = cellValue.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
        } else if (cellValue.length === 14) {
            values[i][0] = cellValue.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, "$1.$2.$3/$4-$5");
        }
    }

    range.setValues(values);
}
