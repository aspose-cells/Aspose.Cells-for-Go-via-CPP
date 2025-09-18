package WorksheetSet

import (
	"fmt"
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Copy Worksheets within a Workbook
func CopyWorksheetsWithinWorkbook() error {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleCopyingAndMovingWorksheets := srcDir + "sampleCopyingAndMovingWorksheets.xlsx"

	// Path of output excel file
	outputCopyingAndMovingWorksheets := outDir + "outputCopyingAndMovingWorksheets.xlsx"

	// Create workbook
	workbook, err := NewWorkbook_String(sampleCopyingAndMovingWorksheets)
	if err != nil {
		return err
	}

	// Create worksheets object with reference to the sheets of the workbook.
	sheets, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}

	// Copy data to a new sheet from an existing sheet within the workbook.
	_, err = sheets.AddCopy_String("Test1")
	if err != nil {
		return err
	}

	// Save the Excel file.
	err = workbook.Save_String(outputCopyingAndMovingWorksheets)
	if err != nil {
		return err
	}

	fmt.Println("Worksheet copied successfully within a workbook!")
	// Show successful execution message on console
	ShowMessageOnConsole("CopyWorksheetsWithInWorkbook executed successfully.\r\n\r\n")
	return nil
}

// Move Worksheets within Workbook
func MoveWorksheetsWithinWorkbook() error {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleCopyingAndMovingWorksheets := srcDir + "sampleCopyingAndMovingWorksheets.xlsx"

	// Path of output excel file
	outputCopyingAndMovingWorksheets := outDir + "outputCopyingAndMovingWorksheets.xlsx"

	// Create workbook
	workbook, err := NewWorkbook_String(sampleCopyingAndMovingWorksheets)
	if err != nil {
		return err
	}

	// Create worksheets object with reference to the sheets of the workbook.
	sheets, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}

	// Access the first sheet
	sheet, err := sheets.Get_Int(0)
	if err != nil {
		return err
	}

	// Move the first sheet to the third position in the workbook.
	err = sheet.MoveTo(2)
	if err != nil {
		return err
	}

	// Save the Excel file.
	err = workbook.Save_String(outputCopyingAndMovingWorksheets)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	ShowMessageOnConsole("MoveWorksheetsWithinWorkbook executed successfully.\r\n\r\n")
	return nil
}
