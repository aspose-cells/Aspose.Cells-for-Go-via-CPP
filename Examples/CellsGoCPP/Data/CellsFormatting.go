package Data

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Format Cell or Range of Cells
func FormatCellOrRangeOfCells() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputFormatCellOrRangeOfCells := outPath + "outputFormatCellOrRangeOfCells.xlsx"

	// Create a new workbook
	wb, _ := NewWorkbook()

	// Get first worksheet which is created by default
	wss, _ := wb.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Access cell C4 by cell name
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("C4")

	// Add some text in cell
	cell.PutValue_String("This is sample data.")

	// Access the cell style
	st, _ := cell.GetStyle()

	// Fill the cell color to Yellow
	st.SetPattern(BackgroundType_Solid)
	color, _ := NewColor()
	color.Set_Color_B(255)
	color.Set_Color_R(255)
	color.Set_Color_G(0)
	st.SetForegroundColor(color)

	// Set the text to wrap
	st.SetIsTextWrapped(true)

	// Set the left and right border to Red
	st.SetBorder_BorderType_CellBorderType_Color(BorderType_LeftBorder, CellBorderType_Thick, color)
	st.SetBorder_BorderType_CellBorderType_Color(BorderType_RightBorder, CellBorderType_Thick, color)

	// Set font color, font size, strike, bold, italic
	font, _ := st.GetFont()
	font.SetColor(color)
	font.SetSize(16)
	font.SetStrikeType(TextStrikeType_Single)
	font.SetIsBold(true)
	font.SetIsItalic(true)

	// Set text horizontal and vertical alignment to center
	st.SetHorizontalAlignment(TextAlignmentType_Center)
	st.SetVerticalAlignment(TextAlignmentType_Center)

	// Set the cell style
	cell.SetStyle_Style(st)

	// Set the cell column width and row height
	column, _ := cell.GetColumn()
	cells.SetColumnWidth(column, 20)
	row, _ := cell.GetRow()
	cells.SetRowHeight(row, 70)

	// Save the output excel file
	wb.Save_String(outputFormatCellOrRangeOfCells)

	// Show successful execution message on console
	fmt.Println("FormatCellOrRangeOfCells executed successfully.\r\n\r\n")
}
