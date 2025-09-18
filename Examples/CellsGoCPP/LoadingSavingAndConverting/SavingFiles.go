package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Saving File to Some Location
func SavingFiletoSomeLocation() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Load sample Excel file
	workbook, _ := NewWorkbook_String(dirPath + "sampleExcelFile.xlsx")

	// Save in Excel 97-2003 format
	workbook.Save_String(outPath + "outputSavingFiletoSomeLocationExcel97-2003.xls")

	// OR
	workbook.Save_String_SaveFormat(outPath+"outputSavingFiletoSomeLocationOrExcel97-2003.xls", SaveFormat_Excel97To2003)

	// Save in Excel2007 xlsx format
	workbook.Save_String_SaveFormat(outPath+"outputSavingFiletoSomeLocationXlsx.xlsx", SaveFormat_Xlsx)

	// Show successful execution message on console
	ShowMessageOnConsole("SavingFiletoSomeLocation executed successfully.\r\n\r\n")
}

// Saving File to a Stream
func SavingFiletoStream() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Load sample Excel file
	workbook, _ := NewWorkbook_String(dirPath + "sampleExcelFile.xlsx")

	// Save the Workbook to Stream
	stream, _ := workbook.Save_SaveFormat(SaveFormat_Xlsx)
	SaveDataToFile(stream, outPath+"outputSavingFiletoStream.xlsx")

	// Show successful execution message on console
	ShowMessageOnConsole("SavingFiletoStream executed successfully.\r\n\r\n")
}
