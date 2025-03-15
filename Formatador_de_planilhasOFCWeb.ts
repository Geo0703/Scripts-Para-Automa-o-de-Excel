function main(workbook: ExcelScript.Workbook) {

    //Format 1   

    let selectedSheet = workbook.getActiveWorksheet();

    // Set vertical alignment to ExcelScript.VerticalAlignment.bottom for range A:A on selectedSheet    

    selectedSheet.getRange("A:A").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

    selectedSheet.getRange("A:A").getFormat().setIndentLevel(0);

    // coloca para baixo    

    selectedSheet.getRange("D:D").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

    selectedSheet.getRange("D:D").getFormat().setIndentLevel(0);

    // coloca para baixo    

    selectedSheet.getRange("E:E").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

    selectedSheet.getRange("E:E").getFormat().setIndentLevel(0);

    // coloca para baixo    

    selectedSheet.getRange("F:F").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

    selectedSheet.getRange("F:F").getFormat().setIndentLevel(0);

    // Add 4 novas colunas    

    selectedSheet.getRange("C:F").insert(ExcelScript.InsertShiftDirection.right);

    //linhas para colunas    

    let rowsToCopy = 4; // Quantas linhas tem que copiar    

    let startRow = 2; // Linha que começa a copiar    

    let startColumn = 1; // Coluna de origem    

    let destRow = 4; // Linha de destino (onde queremos colar os resultados -1)    

    let destColumn = 2; // para onde vai    

    // Loop para copiar    

    for (let i = 0; i < 1000; i++) { // loop de 1000×    

        let sourceRange = selectedSheet.getRangeByIndexes(startRow + i * rowsToCopy - 1, startColumn, rowsToCopy, 1);

        let destRange = selectedSheet.getRangeByIndexes(destRow + i * rowsToCopy, destColumn, 1, 4); // 4 colunas para as 4 copias    

        destRange.copyFrom(sourceRange, ExcelScript.RangeCopyType.all, false, true);

    }

    //Format 2    

    // formata para o tamanho    

    selectedSheet.getRange("C:C").getFormat().autofitColumns();

    // formata para o tamanho  

    selectedSheet.getRange("D:D").getFormat().autofitColumns();

    // formata para o tamanho    

    selectedSheet.getRange("E:E").getFormat().autofitColumns();

    // formata para o tamanho    

    selectedSheet.getRange("F:F").getFormat().autofitColumns();

    // Deixa B oculta    

    selectedSheet.getRange("B:B").delete(ExcelScript.DeleteShiftDirection.left);

    //tira mesclagem

    selectedSheet.getRange("A:A").unmerge();

    selectedSheet.getRange("G:G").unmerge();

    selectedSheet.getRange("H:H").unmerge();

    selectedSheet.getRange("I:I").unmerge();

    //nomeia as novas colunas

    selectedSheet.getRange("B1:E1").setValues([["Processo", "Contrato", "Credor", "Razão"]])

    selectedSheet.getRange("G2:I10000").moveTo(selectedSheet.getRange("G5"));

    selectedSheet.getRange("A2:A10000").moveTo(selectedSheet.getRange("A5"));

    selectedSheet.getRange("F4:F10000").moveTo(selectedSheet.getRange("F5"));

    let sheet = workbook.getActiveWorksheet();
    let totalRows = sheet.getUsedRange().getRowCount();

    // Definir o padrão para a exclusão de linhas
    let deleteCount = 3; // Número de linhas para excluir
    let skipCount = 1;   // Número de linhas para pular

    // Começa a partir da linha 2 e move-se para baixo
    let currentRow = 2;  // Começar a partir da linha 2

    // Percorrer até o final da planilha
    while (currentRow <= totalRows) {
        // Excluir linhas
        for (let i = 0; i < deleteCount; i++) {
            if (currentRow > totalRows) break; // Sai se não houver mais linhas
            sheet.getRange(`A${currentRow}`).getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
            // Não aumentar currentRow aqui, pois a linha abaixo subirá após a exclusão
        }
        // Pular linhas
        for (let i = 0; i < skipCount; i++) {
            currentRow++;  // Aumentar currentRow para pular as linhas
            if (currentRow > totalRows) break; // Sai se não houver mais linhas
        }
    }
}

