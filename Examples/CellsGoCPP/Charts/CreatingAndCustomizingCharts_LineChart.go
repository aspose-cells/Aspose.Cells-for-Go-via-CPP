package Charts

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Creating and Customizing Charts - Line Chart
func CreatingAndCustomizingCharts_LineChart() {
	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output excel file
	outputChartTypeLine := outDir + "outputChartTypeLine.xlsx"

	// Create a new workbook
	workbook, _ := NewWorkbook()

	// Get first worksheet which is created by default
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Adding sample values to cells
	cells, _ := worksheet.GetCells()
	cell, _ := cells.Get_String("A1")
	cell.PutValue_Int(50)
	cell, _ = cells.Get_String("A2")
	cell.PutValue_Int(100)
	cell, _ = cells.Get_String("A3")
	cell.PutValue_Int(150)
	cell, _ = cells.Get_String("B1")
	cell.PutValue_Int(4)
	cell, _ = cells.Get_String("B2")
	cell.PutValue_Int(20)
	cell, _ = cells.Get_String("B3")
	cell.PutValue_Int(50)

	// Adding a chart to the worksheet
	charts, _ := worksheet.GetCharts()
	chartIndex, _ := charts.Add_ChartType_Int_Int_Int_Int(ChartType_Line, 5, 0, 20, 8)
	// Accessing the instance of the newly added chart
	charts, _ = worksheet.GetCharts()
	chart, _ := charts.Get_Int(chartIndex)

	// Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
	nSeries, _ := chart.GetNSeries()
	nSeries.Add_String_Bool("A1:B3", true)

	// Saving the Excel file
	workbook.Save_String(outputChartTypeLine)

	// Show successful execution message on console
	fmt.Println("CreatingAndCustomizingCharts_LineChart executed successfully.")
}
