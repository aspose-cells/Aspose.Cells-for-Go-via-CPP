package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Opening Excel File using its Path
func OpeningExcelFileUsingPath() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// Create Workbook object from an Excel file path
	NewWorkbook_String(dirPath + "sampleExcelFile.xlsx")

	// Show following message on console
	ShowMessageOnConsole("Workbook opened successfully using file path.")

	// Show successful execution message on console
	ShowMessageOnConsole("OpeningExcelFileUsingPath executed successfully.\r\n\r\n")
}

// Opening Excel File using Stream
func OpeningExcelFileUsingStream() {
	// Source directory path
	dirPath := "..\\Data\\LoadingSavingAndConverting\\"

	// You need to write your own code to read the contents of the sampleExcelFile.xlsx file into this variable.
	fileStream, _ := GetDataFromFile(dirPath + "sampleExcelFile.xlsx")

	// Create Workbook object from a Stream object
	NewWorkbook_Stream(fileStream)

	// Show following message on console
	ShowMessageOnConsole("Workbook opened successfully using stream.")

	// Show successful execution message on console
	ShowMessageOnConsole("OpeningExcelFileUsingStream executed successfully.\r\n\r\n")
}
