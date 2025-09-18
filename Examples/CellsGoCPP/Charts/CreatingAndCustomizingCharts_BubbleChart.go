package Charts

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// Creating and Customizing Charts - Bubble Chart
func CreatingAndCustomizingCharts_BubbleChart() {
	// Path of output excel file
	outputChartTypeBubble := "outputChartTypeBubble.xlsx"

	// Create a new workbook
	workbook, _ := NewWorkbook()

	// Get first worksheet which is created by default
	wss, _ := workbook.GetWorksheets()
	ws, _ := wss.Get_Int(0)

	// Fill in data for chart's series
	// Y Values
	cells, _ := ws.GetCells()
	cell, _ := cells.Get_String("A1")
	cell.PutValue_String("Y Values")
	cell, _ = cells.Get_String("B1")
	cell.PutValue_Int(2)
	cell, _ = cells.Get_String("C1")
	cell.PutValue_Int(4)
	cell, _ = cells.Get_String("D1")
	cell.PutValue_Int(6)

	// Bubble Size
	cell, _ = cells.Get_String("A2")
	cell.PutValue_String("Bubble Size")
	cell, _ = cells.Get_String("B2")
	cell.PutValue_Int(2)
	cell, _ = cells.Get_String("C2")
	cell.PutValue_Int(3)
	cell, _ = cells.Get_String("D2")
	cell.PutValue_Int(1)

	// X Values
	cell, _ = cells.Get_String("A3")
	cell.PutValue_String("X Values")
	cell, _ = cells.Get_String("B3")
	cell.PutValue_Int(1)
	cell, _ = cells.Get_String("C3")
	cell.PutValue_Int(2)
	cell, _ = cells.Get_String("D3")
	cell.PutValue_Int(3)

	// Set first column width
	cells.SetColumnWidth(0, 12)

	// Adding a chart to the worksheet
	charts, _ := ws.GetCharts()
	chartIndex, _ := charts.Add_ChartType_Int_Int_Int_Int(ChartType_Bubble, 5, 0, 20, 8)

	// Accessing the instance of the newly added chart
	charts, _ = ws.GetCharts()
	chart, _ := charts.Get_Int(chartIndex)

	// Adding SeriesCollection (chart data source) to the chart ranging from B1 to D1
	nSeries, _ := chart.GetNSeries()
	nSeries.Add_String_Bool("B1:D1", true)

	// Set bubble sizes1
	nSeries, _ = chart.GetNSeries()
	series, _ := nSeries.Get(0)
	series.SetBubbleSizes("B2:D2")

	// Set X axis values
	series.SetXValues("B3:D3")

	// Set Y axis values
	series.SetValues("B1:D1")

	// Saving the Excel file
	workbook.Save_String(outputChartTypeBubble)

	// Show successful execution message on console
	fmt.Println("CreatingAndCustomizingCharts_BubbleChart executed successfully.")
}
