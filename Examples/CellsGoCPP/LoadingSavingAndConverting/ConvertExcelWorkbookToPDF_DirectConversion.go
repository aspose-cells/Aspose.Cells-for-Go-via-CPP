package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ConvertExcelWorkbookToPDF_DirectConversion converts an Excel Workbook to PDF - Direct Conversion
func ConvertExcelWorkbookToPDF_DirectConversion() {
	// Source directory path.
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path.
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input Excel file
	sampleConvertExcelWorkbookToPDF := srcDir + "sampleConvertExcelWorkbookToPDF.xlsx"

	// Path of output Pdf file
	outputConvertExcelWorkbookToPDF := outDir + "outputConvertExcelWorkbookToPDF_DirectConversion.pdf"

	// Load the sample Excel file.
	workbook, _ := NewWorkbook_String(sampleConvertExcelWorkbookToPDF)

	// Save the Excel Document in PDF format
	err := workbook.Save_String_SaveFormat(outputConvertExcelWorkbookToPDF, SaveFormat_Pdf)
	if err != nil {
		fmt.Println("Error saving workbook to PDF:", err)
		return
	}

	// Show successful execution message on console
	fmt.Println("ConvertExcelWorkbookToPDF_DirectConversion executed successfully.")
}
