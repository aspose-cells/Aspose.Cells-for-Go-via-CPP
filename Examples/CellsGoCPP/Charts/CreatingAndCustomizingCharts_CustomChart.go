package Charts

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

func CreatingAndCustomizingCharts_CustomChart() error {
	// Output directory path
	outDir := "..\\Data\\02_OutputDirectory\\"

	// Path of output excel file
	outputChartTypeCustom := outDir + "outputChartTypeCustom.xlsx"

	// Create a new workbook
	workbook, err := NewWorkbook()
	if err != nil {
		return err
	}

	// Get first worksheet which is created by default
	wss, err := workbook.GetWorksheets()
	if err != nil {
		return err
	}
	worksheet, err := wss.Get_Int(0)
	if err != nil {
		return err
	}

	// Adding sample values to cells
	cells, err := worksheet.GetCells()
	if err != nil {
		return err
	}

	cell, err := cells.Get_String("A1")
	if err != nil {
		return err
	}
	cell.PutValue_Int(50)

	cell, err = cells.Get_String("A2")
	if err != nil {
		return err
	}
	cell.PutValue_Int(100)

	cell, err = cells.Get_String("A3")
	if err != nil {
		return err
	}
	cell.PutValue_Int(150)

	cell, err = cells.Get_String("A4")
	if err != nil {
		return err
	}
	cell.PutValue_Int(110)

	cell, err = cells.Get_String("B1")
	if err != nil {
		return err
	}
	cell.PutValue_Int(260)

	cell, err = cells.Get_String("B2")
	if err != nil {
		return err
	}
	cell.PutValue_Int(12)

	cell, err = cells.Get_String("B3")
	if err != nil {
		return err
	}
	cell.PutValue_Int(50)

	cell, err = cells.Get_String("B4")
	if err != nil {
		return err
	}
	cell.PutValue_Int(100)

	// Adding a chart to the worksheet
	charts, err := worksheet.GetCharts()
	chartIndex, err := charts.Add_ChartType_Int_Int_Int_Int(ChartType_Column, 5, 0, 20, 8)
	if err != nil {
		return err
	}

	// Accessing the instance of the newly added chart
	charts, err = worksheet.GetCharts()
	if err != nil {
		return err
	}
	chart, err := charts.Get_Int(chartIndex)
	if err != nil {
		return err
	}

	// Adding SeriesCollection (chart data source) to the chart ranging from A1 to B4
	nSeries, err := chart.GetNSeries()
	_, err = nSeries.Add_String_Bool("A1:B4", true)
	if err != nil {
		return err
	}

	// Setting the chart type of 2nd NSeries to display as line chart
	nSeries, err = chart.GetNSeries()
	if err != nil {
		return err
	}
	series, err := nSeries.Get(1)
	if err != nil {
		return err
	}
	series.SetType(ChartType_Line)

	// Saving the Excel file
	err = workbook.Save_String(outputChartTypeCustom)
	if err != nil {
		return err
	}

	// Show successful execution message on console
	fmt.Println("CreatingAndCustomizingCharts_CustomChart executed successfully.")
	return nil
}
