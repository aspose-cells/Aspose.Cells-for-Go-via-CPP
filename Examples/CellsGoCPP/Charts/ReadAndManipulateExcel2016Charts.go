package Charts

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

func ReadAndManipulateExcel2016Charts() {
	// Source directory path
	srcDir := "..\\Data\\01_SourceDirectory\\"

	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of input excel file
	sampleReadAndManipulateExcel2016Charts := srcDir + "sampleReadAndManipulateExcel2016Charts.xlsx"

	// Path of output excel file
	outputReadAndManipulateExcel2016Charts := outDir + "outputReadAndManipulateExcel2016Charts.xlsx"

	// Load sample Excel file containing Excel 2016 charts
	workbook, _ := NewWorkbook_String(sampleReadAndManipulateExcel2016Charts)

	// Access the first worksheet which contains the charts
	wss, _ := workbook.GetWorksheets()
	worksheet, _ := wss.Get_Int(0)

	// Access all charts one by one and read their types
	charts, _ := worksheet.GetCharts()
	count, _ := charts.GetCount()
	var i int32
	for i = 0; i < count; i++ {
		// Access the chart
		charts, _ := worksheet.GetCharts()
		ch, _ := charts.Get_Int(i)

		// Get the chart type
		chartType, _ := ch.GetType()

		// Convert chart type enum to string
		var strChartType string

		switch chartType {
		case ChartType_BoxWhisker:
			strChartType = "BoxWhisker"
		case ChartType_Histogram:
			strChartType = "Histogram"
		case ChartType_Sunburst:
			strChartType = "Sunburst"
		case ChartType_Treemap:
			strChartType = "Treemap"
		case ChartType_Waterfall:
			strChartType = "Waterfall"
		default:
			continue
		}

		// Print chart type
		fmt.Println(strChartType)

		// Change the title of the charts as per their types
		strTitle := "Chart Type is " + strChartType
		title, _ := ch.GetTitle()
		title.SetText(strTitle)
	}

	// Save the workbook
	workbook.Save_String(outputReadAndManipulateExcel2016Charts)

	// Show successful execution message on console
	fmt.Println("ReadAndManipulateExcel2016Charts executed successfully.")
}
