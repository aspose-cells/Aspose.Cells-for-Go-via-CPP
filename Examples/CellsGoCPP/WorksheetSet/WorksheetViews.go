package WorksheetSet

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// EnablingPageBreakPreview
func EnablingPageBreakPreview() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleWorksheetViews := srcDir + "sampleWorksheetViews.xlsx"

	// Path of input excel file
	outputWorksheetViews := outDir + "outputWorksheetViews.xlsx"

	// Instantiate a workbook object
	workbook, _ := NewWorkbook_String(sampleWorksheetViews)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Displaying the worksheet in page break preview
	worksheet.SetIsPageBreakPreview(true)

	// Save the Excel file
	workbook.Save_String(outputWorksheetViews)

	// Show successful execution message on console
	ShowMessageOnConsole("EnablingPageBreakPreview executed successfully.\r\n\r\n")
}

// ZoomFactor
func ZoomFactor() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleWorksheetViews := srcDir + "sampleWorksheetViews.xlsx"

	// Path of input excel file
	outputWorksheetViews := outDir + "outputWorksheetViews.xlsx"

	// Instantiate a workbook object
	workbook, _ := NewWorkbook_String(sampleWorksheetViews)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Setting the zoom factor of the worksheet to 75
	worksheet.SetZoom(75)

	// Saving the modified Excel file
	workbook.Save_String(outputWorksheetViews)

	// Show successful execution message on console
	ShowMessageOnConsole("ZoomFactor executed successfully.\r\n\r\n")
}

// FreezePanes
func FreezePanes() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleWorksheetViews := srcDir + "sampleWorksheetViews.xlsx"

	// Path of input excel file
	outputWorksheetViews := outDir + "outputWorksheetViews.xlsx"

	// Instantiating a Workbook object and opening the Excel file through the file stream
	workbook, _ := NewWorkbook_String(sampleWorksheetViews)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Applying freeze panes settings
	worksheet.FreezePanes_Int_Int_Int_Int(3, 2, 3, 2)

	// Saving the modified Excel file
	workbook.Save_String(outputWorksheetViews)

	// Show successful execution message on console
	ShowMessageOnConsole("FreezePanes executed successfully.\r\n\r\n")
}

// SplitPanes
func SplitPanes() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleWorksheetViews := srcDir + "sampleWorksheetViews.xlsx"

	// Path of input excel file
	outputWorksheetViews := outDir + "outputWorksheetViews.xlsx"

	// Instantiating a Workbook object
	workbook, _ := NewWorkbook_String(sampleWorksheetViews)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Set the active cell
	worksheet.SetActiveCell("A20")

	// Split the worksheet window
	worksheet.Split()

	// Saving the modified Excel file
	workbook.Save_String(outputWorksheetViews)

	// Show successful execution message on console
	ShowMessageOnConsole("SplitPanes executed successfully.\r\n\r\n")
}

// RemovingPanes
func RemovingPanes() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleWorksheetViews := srcDir + "sampleWorksheetViews.xlsx"

	// Path of input excel file
	outputWorksheetViews := outDir + "outputWorksheetViews.xlsx"

	// Instantiating a Workbook object
	workbook, _ := NewWorkbook_String(sampleWorksheetViews)

	// Accessing a worksheet using its index
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Set the active cell
	worksheet.SetActiveCell("A20")

	// Split the worksheet window
	worksheet.RemoveSplit()

	// Saving the modified Excel file
	workbook.Save_String(outputWorksheetViews)

	// Show successful execution message on console
	ShowMessageOnConsole("RemovingPanes executed successfully.\r\n\r\n")
}
