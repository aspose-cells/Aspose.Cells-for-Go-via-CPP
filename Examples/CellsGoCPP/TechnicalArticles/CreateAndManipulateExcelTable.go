package TechnicalArticles

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// CreateAndManipulateExcelTable creates and manipulates an Excel table.
func CreateAndManipulateExcelTable() error {
	// Source directory path
	dirPath := "..\\Data\\TechnicalArticles\\"

	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of input excel file
	sampleCreateAndManipulateExcelTable := dirPath + "sampleCreateAndManipulateExcelTable.xlsx"

	// Path of output excel file
	outputCreateAndManipulateExcelTable := outPath + "outputCreateAndManipulateExcelTable.xlsx"

	// Load the sample excel file
	wb, err := NewWorkbook_String(sampleCreateAndManipulateExcelTable)
	if err != nil {
		return err
	}

	// Access first worksheet
	wss, err := wb.GetWorksheets()
	if err != nil {
		return err
	}
	ws, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Add table i.e., list object
	listObjects, err := ws.GetListObjects()
	idx, err := listObjects.Add_String_String_Bool("A1", "H10", true)
	if err != nil {
		return err
	}

	// Access the newly added list object
	listObjects, err = ws.GetListObjects()
	lo, err := listObjects.Get_Int(idx)
	if err != nil {
		return err
	}

	// Use its display methods
	err = lo.SetShowHeaderRow(true)
	if err != nil {
		return err
	}
	err = lo.SetShowTableStyleColumnStripes(true)
	if err != nil {
		return err
	}
	err = lo.SetShowTotals(true)
	if err != nil {
		return err
	}

	// Set its style
	err = lo.SetTableStyleType(TableStyleType_TableStyleLight12)
	if err != nil {
		return err
	}

	// Set total functions of 3rd, 4th and 5th columns
	columns, err := lo.GetListColumns()
	if err != nil {
		return err
	}
	col3, err := columns.Get_Int(2)
	if err != nil {
		return err
	}
	err = col3.SetTotalsCalculation(TotalsCalculation_Min)
	if err != nil {
		return err
	}
	col4, err := columns.Get_Int(3)
	if err != nil {
		return err
	}
	err = col4.SetTotalsCalculation(TotalsCalculation_Max)
	if err != nil {
		return err
	}
	col5, err := columns.Get_Int(4)
	if err != nil {
		return err
	}
	err = col5.SetTotalsCalculation(TotalsCalculation_Count)
	if err != nil {
		return err
	}

	// Save the output excel file
	err = wb.Save_String(outputCreateAndManipulateExcelTable)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("CreateAndManipulateExcelTable executed successfully.\r\n\r\n")

	return nil
}
