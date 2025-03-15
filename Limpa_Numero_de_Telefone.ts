//Limpa Numero de telefone
function main(workbook: ExcelScript.Workbook) {
    // Obtenha a planilha ativa
    let sheet = workbook.getActiveWorksheet();
    let selectedSheet = workbook.getActiveWorksheet();

    // Set format for range A:A on selectedSheet
    selectedSheet.getRange("B:B").setNumberFormatLocal("1");

    // Obtenha todas as células na coluna A
    let range = sheet.getRange("A1:A" + sheet.getUsedRange().getRowCount());
    let values = range.getValues();

    // Itere sobre cada célula na coluna A
    for (let i = 0; i < values.length; i++) {
        let cellValue = values[i][0].toString();

        // Verifique se o primeiro número é 9
        if (cellValue.startsWith("9")) {
            cellValue = "5511" + cellValue;
        }

        // Verifique se o primeiro número é 11
        else if (cellValue.startsWith("11")) {
            cellValue = "55" + cellValue;
        }

        // Remova traços, espaços e parênteses
        cellValue = cellValue.replace(/[-\s()]/g, "");

        // Atualize o valor da célula
        values[i][0] = cellValue;
    }

    // Defina os valores atualizados de volta na coluna A
    range.setValues(values);
}
