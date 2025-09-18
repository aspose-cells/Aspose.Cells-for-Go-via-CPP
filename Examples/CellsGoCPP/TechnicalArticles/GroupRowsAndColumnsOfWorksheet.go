package TechnicalArticles

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// GroupRowsAndColumnsOfWorksheet groups rows and columns of a worksheet
func GroupRowsAndColumnsOfWorksheet() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputGroupRowsAndColumnsOfWorksheet := outPath + "outputGroupRowsAndColumnsOfWorksheet.xlsx"

	// Create an empty workbook
	wb, _ := NewWorkbook()

	// Add worksheet for grouping rows
	wss, _ := wb.GetWorksheets()
	grpRows, _ := wss.Get_Int(0)
	grpRows.SetName("GroupRows")

	// Add worksheet for grouping columns
	idx, _ := wss.Add()
	grpCols, _ := wss.Get_Int(idx)
	grpCols.SetName("GroupColumns")

	// Add sample values in both worksheets
	var i int32
	for i = 0; i < 50; i++ {
		str := "Text"

		cells1, _ := grpRows.GetCells()
		cell1, _ := cells1.Get_Int_Int(i, 0)
		cell1.PutValue_String(str)

		cells2, _ := grpCols.GetCells()
		cell2, _ := cells2.Get_Int_Int(0, i)
		cell2.PutValue_String(str)
	}

	// Grouping rows at first level
	cellsGrpRows, _ := grpRows.GetCells()
	cellsGrpRows.GroupRows_Int_Int(0, 10)
	cellsGrpRows.GroupRows_Int_Int(12, 22)
	cellsGrpRows.GroupRows_Int_Int(24, 34)

	// Grouping rows at second level
	cellsGrpRows.GroupRows_Int_Int(2, 8)
	cellsGrpRows.GroupRows_Int_Int(14, 20)
	cellsGrpRows.GroupRows_Int_Int(28, 30)

	// Grouping rows at third level
	cellsGrpRows.GroupRows_Int_Int(5, 7)

	// Grouping columns at first level
	cellsGrpCols, _ := grpCols.GetCells()
	cellsGrpCols.GroupColumns_Int_Int(0, 10)
	cellsGrpCols.GroupColumns_Int_Int(12, 22)
	cellsGrpCols.GroupColumns_Int_Int(24, 34)

	// Grouping columns at second level
	cellsGrpCols.GroupColumns_Int_Int(2, 8)
	cellsGrpCols.GroupColumns_Int_Int(14, 20)
	cellsGrpCols.GroupColumns_Int_Int(28, 30)

	// Grouping columns at third level
	cellsGrpCols.GroupColumns_Int_Int(5, 7)

	// Save the output excel file
	wb.Save_String(outputGroupRowsAndColumnsOfWorksheet)

	// Show successful execution message on console
	fmt.Println("GroupRowsAndColumnsOfWorksheet executed successfully.\r\n\r\n")
}
