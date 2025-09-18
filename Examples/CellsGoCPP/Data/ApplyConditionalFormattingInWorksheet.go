package Data

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ApplyConditionalFormattingInWorksheet applies conditional formatting in a worksheet
func ApplyConditionalFormattingInWorksheet() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputApplyConditionalFormattingInWorksheet := outPath + "outputApplyConditionalFormattingInWorksheet.xlsx"

	// Create an empty workbook
	wb, _ := NewWorkbook()

	// Access first worksheet
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Adds an empty conditional formatting
	cfs, _ := ws.GetConditionalFormattings()
	idx, _ := cfs.Add()
	fcs, _ := cfs.Get(idx)

	// Set the conditional format range
	ca, _ := CellArea_CreateCellArea_String_String("A1", "A1")
	fcs.AddArea(ca)
	ca, _ = CellArea_CreateCellArea_String_String("B2", "B2")
	fcs.AddArea(ca)

	// Add condition and set the background color
	idx, _ = fcs.AddCondition_FormatConditionType_OperatorType_String_String(FormatConditionType_CellValue, OperatorType_Between, "=A2", "100")
	fc, _ := fcs.Get(idx)
	style, _ := fc.GetStyle()
	color, _ := NewColor()
	color.Set_Color_B(255)
	color.Set_Color_R(255)
	color.Set_Color_G(0)
	style.SetBackgroundColor(color)

	// User friendly message to test the output excel file
	msgStr := "Red color in cells A1 and B2 is because of Conditional Formatting. Put 101 or any value >100 in cell A2 and B2, you will see Red background color will be gone."
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("A10")
	cell.PutValue_String(msgStr)

	// Save the output excel file
	wb.Save_String(outputApplyConditionalFormattingInWorksheet)

	// Show successful execution message on console
	fmt.Println("ApplyConditionalFormattingInWorksheet executed successfully.\r\n\r\n")
}
