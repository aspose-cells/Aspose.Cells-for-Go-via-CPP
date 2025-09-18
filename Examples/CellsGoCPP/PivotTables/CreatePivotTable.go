package LoadingSavingAndConverting

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// CreatePivotTable creates a Pivot Table
func CreatePivotTable() error {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputCreatePivotTable := outPath + "outputCreatePivotTable.xlsx"

	// Load the sample excel file
	wb, err := NewWorkbook()
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

	// Add source data for pivot table
	str := "Fruit"
	cells, err := ws.GetCells()
	if err != nil {
		return err
	}
	cell, err := cells.Get_String("A1")
	if err != nil {
		return err
	}
	cell.PutValue_String(str)

	str = "Quantity"
	cell, err = cells.Get_String("B1")
	if err != nil {
		return err
	}
	cell.PutValue_String(str)

	str = "Price"
	cell, err = cells.Get_String("C1")
	if err != nil {
		return err
	}
	cell.PutValue_String(str)

	str = "Apple"
	cell, err = cells.Get_String("A2")
	if err != nil {
		return err
	}
	cell.PutValue_String(str)

	str = "Orange"
	cell, err = cells.Get_String("A3")
	if err != nil {
		return err
	}
	cell.PutValue_String(str)

	cell, err = cells.Get_String("B2")
	if err != nil {
		return err
	}
	cell.PutValue_Int(3)

	cell, err = cells.Get_String("B3")
	if err != nil {
		return err
	}
	cell.PutValue_Int(4)

	cell, err = cells.Get_String("C2")
	if err != nil {
		return err
	}
	cell.PutValue_Int(2)

	cell, err = cells.Get_String("C3")
	if err != nil {
		return err
	}
	cell.PutValue_Int(1)

	// Add pivot table
	pivotTables, err := ws.GetPivotTables()
	idx, err := pivotTables.Add_String_String_String("A1:C3", "E5", "MyPivotTable")
	if err != nil {
		return err
	}

	// Access created pivot table
	pivotTables, err = ws.GetPivotTables()
	if err != nil {
		return err
	}
	pt, err := pivotTables.Get_Int(idx)
	if err != nil {
		return err
	}

	// Manipulate pivot table rows, columns and data fields
	baseFields, err := pt.GetBaseFields()
	if err != nil {
		return err
	}
	rowField, err := baseFields.Get_Int(0)
	if err != nil {
		return err
	}
	pt.AddFieldToArea_PivotFieldType_PivotField(PivotFieldType_Row, rowField)

	dataField1, err := baseFields.Get_Int(1)
	if err != nil {
		return err
	}
	pt.AddFieldToArea_PivotFieldType_PivotField(PivotFieldType_Data, dataField1)

	dataField2, err := baseFields.Get_Int(2)
	if err != nil {
		return err
	}
	pt.AddFieldToArea_PivotFieldType_PivotField(PivotFieldType_Data, dataField2)

	dataField, err := pt.GetDataField()
	if err != nil {
		return err
	}
	pt.AddFieldToArea_PivotFieldType_PivotField(PivotFieldType_Column, dataField)

	// Set pivot table style
	pt.SetPivotTableStyleType(PivotTableStyleType_PivotTableStyleMedium9)

	// Save the output excel file
	err = wb.Save_String(outputCreatePivotTable)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("CreatePivotTable executed successfully.\r\n\r\n")
	return nil
}
