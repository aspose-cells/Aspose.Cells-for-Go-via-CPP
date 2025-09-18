package LoadingSavingAndConverting

import (
	. "main/Common"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Grouping Rows & Columns
func GroupingRowsColumns() error {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleGroupingUngroupingRowsAndColumns := dirPath + "sampleGroupingUngroupingRowsAndColumns.xlsx"

	// Path of output excel file
	outputGroupingUngroupingRowsAndColumns := outPath + "outputGroupingUngroupingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, err := NewWorkbook_String(sampleGroupingUngroupingRowsAndColumns)
	if err != nil {
		return err
	}

	// Accessing the first worksheet in the Excel file
	wss, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	worksheet, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Grouping first seven rows and first four columns
	cells, err := worksheet.GetCells()
	if err != nil {
		return err
	}
	err = cells.GroupRows_Int_Int_Bool(0, 6, true)
	if err != nil {
		return err
	}
	err = cells.GroupColumns_Int_Int_Bool(0, 3, true)
	if err != nil {
		return err
	}

	// Save the Excel file.
	err = workbook.Save_String(outputGroupingUngroupingRowsAndColumns)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	ShowMessageOnConsole("GroupingRowsColumns executed successfully.\r\n\r\n")
	return nil
}

// Ungrouping Rows & Columns
func UnGroupingRowsColumns() error {
	// Source directory path
	dirPath := "..\\Data\\RowsAndColumns\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleGroupingUngroupingRowsAndColumns := dirPath + "sampleGroupingUngroupingRowsAndColumns.xlsx"

	// Path of output excel file
	outputGroupingUngroupingRowsAndColumns := outPath + "outputGroupingUngroupingRowsAndColumns.xlsx"

	// Read input excel file
	workbook, err := NewWorkbook_String(sampleGroupingUngroupingRowsAndColumns)
	if err != nil {
		return err
	}

	// Accessing the second worksheet in the Excel file
	wss, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	worksheet, err := wss.Get_Int(1)
	if err != nil {
		return err
	}

	// UnGroup first seven rows and first four columns
	cells, err := worksheet.GetCells()
	if err != nil {
		return err
	}
	err = cells.UngroupRows_Int_Int(0, 6)
	if err != nil {
		return err
	}
	err = cells.UngroupColumns(0, 3)
	if err != nil {
		return err
	}

	// Save the Excel file.
	err = workbook.Save_String(outputGroupingUngroupingRowsAndColumns)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	ShowMessageOnConsole("UnGroupingRowsColumns executed successfully.\r\n\r\n")
	return nil
}
