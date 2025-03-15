function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Set width of column(s) at range A:XFD on selectedSheet to 219
	selectedSheet.getRange("A:XFD").getFormat().setColumnWidth(219);
	// Auto fit the columns of all cells on selectedSheet
	selectedSheet.getRange().getFormat().autofitColumns();
	// Set fill color to FFFFFF for range 1:1 on selectedSheet
	selectedSheet.getRange("1:1").getFormat().getFill().setColor("FFFFFF");
	// Set font color to "000000" for range 1:1 on selectedSheet
	selectedSheet.getRange("1:1").getFormat().getFont().setColor("000000");
	// Set height of row(s) at range 1:1 on selectedSheet to 29.25
	selectedSheet.getRange("1:1").getFormat().setRowHeight(29.25);
	// Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range 1:1 on selectedSheet
	selectedSheet.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	selectedSheet.getRange("1:1").getFormat().setIndentLevel(0);
	// Set vertical alignment to ExcelScript.VerticalAlignment.center for range 1:1 on selectedSheet
	selectedSheet.getRange("1:1").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	selectedSheet.getRange("1:1").getFormat().setIndentLevel(0);
}
