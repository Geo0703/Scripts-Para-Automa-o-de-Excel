function main(workbook: ExcelScript.Workbook) {
	// Adicionar uma nova planilha com o nome "Geral"
	let geral = workbook.addWorksheet("Geral1");

	// Definir as planilhas
	let gEX_ABCD = workbook.getWorksheet("GEX ABCD");
	let gEX_ARA_ATUBA = workbook.getWorksheet("GEX ARAÇATUBA");
	let gEX_ARARAQUARA = workbook.getWorksheet("GEX ARARAQUARA");
	let gEX_BAURU = workbook.getWorksheet("GEX BAURU");
	let gEXCPN = workbook.getWorksheet("GEXCPN");
	let gEXGRU = workbook.getWorksheet("GEXGRU");
	let gEXJDI = workbook.getWorksheet("GEXJDI");
	let gEXMRI = workbook.getWorksheet("GEXMRI");
	let gEXOSA = workbook.getWorksheet("GEXOSA");
	let gEXPIR = workbook.getWorksheet("GEXPIR");
	let gEXPRP = workbook.getWorksheet("GEXPRP");
	let gEXRBP = workbook.getWorksheet("GEXRBP");
	let gEXSAN = workbook.getWorksheet("GEXSAN");
	let gEXSRP = workbook.getWorksheet("GEXSRP");
	let gEXSBV = workbook.getWorksheet("GEXSBV");
	let gEXSOR = workbook.getWorksheet("GEXSOR");
	let gEXSP = workbook.getWorksheet("GEXSP");
	let gEXVPB = workbook.getWorksheet("GEXVPB");

	// Colar dados de cada planilha na "Geral"
	geral.getRange("A1").copyFrom(gEX_ABCD.getRange("A1:BE22"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A23").copyFrom(gEX_ARA_ATUBA.getRange("A2:BE40"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A62").copyFrom(gEX_ARARAQUARA.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A122").copyFrom(gEX_BAURU.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A173").copyFrom(gEXCPN.getRange("A2:BE46"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A218").copyFrom(gEXGRU.getRange("A2:BE28"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A245").copyFrom(gEXJDI.getRange("A2:BE31"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A275").copyFrom(gEXMRI.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A320").copyFrom(gEXOSA.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A347").copyFrom(gEXPIR.getRange("A2:BD40"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A386").copyFrom(gEXPRP.getRange("A2:BD2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A422").copyFrom(gEXRBP.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A467").copyFrom(gEXSAN.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A515").copyFrom(gEXSRP.getRange("A2:BE55"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A569").copyFrom(gEXSBV.getRange("A2:BE49"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A617").copyFrom(gEXSOR.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A680").copyFrom(gEXSP.getRange("A2:BE116"), ExcelScript.RangeCopyType.all, false, false);
	geral.getRange("A795").copyFrom(gEXVPB.getRange("A2:BE2").getExtendedRange(ExcelScript.KeyboardDirection.down), ExcelScript.RangeCopyType.all, false, false);

	// Definir a largura das colunas A:XFD na "Geral" para 564
	geral.getRange("A:XFD").getFormat().setColumnWidth(564);
	// Ajustar automaticamente a largura das colunas de todas as células na "Geral"
	geral.getRange().getFormat().autofitColumns();
	// Definir a altura das linhas na "Geral" para 67.5
	geral.getRange("A:XFD").getFormat().setRowHeight(67.5);
	// Ajustar automaticamente a altura das linhas de todas as células na "Geral"
	geral.getRange().getFormat().autofitRows();
}
