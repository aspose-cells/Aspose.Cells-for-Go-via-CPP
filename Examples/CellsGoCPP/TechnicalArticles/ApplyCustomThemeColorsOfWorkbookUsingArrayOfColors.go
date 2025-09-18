package TechnicalArticles

import (
	"fmt"

	. "github.com/aspose-cells/aspose-cells-go-cpp/v25"
)

// ApplyCustomThemeColorsOfWorkbookUsingArrayOfColors applies custom theme colors to the workbook using an array of colors.
func ApplyCustomThemeColorsOfWorkbookUsingArrayOfColors() {
	// Output directory path
	outPath := "..\\Data\\Output\\"

	// Path of output excel file
	outputApplyCustomThemeColorsOfWorkbookUsingArrayOfColors := outPath + "outputApplyCustomThemeColorsOfWorkbookUsingArrayOfColors.xlsx"

	// Create a workbook
	wb, _ := NewWorkbook()

	// Create array of custom theme colors
	clrs := make([]Color, 12)
	// Background1
	color, _ := NewColor()
	color.Set_Color_B(255)
	color.Set_Color_R(255)
	color.Set_Color_G(0)
	clrs[0] = *color
	// Text1
	clrs[1] = *color
	// Background2
	clrs[2] = *color
	// Text2
	clrs[3] = *color
	// Accent1
	clrs[4] = *color

	color, _ = NewColor()
	color.Set_Color_B(255)
	color.Set_Color_R(0)
	color.Set_Color_G(255)
	// Accent2
	clrs[5] = *color
	// Accent3
	clrs[6] = *color
	// Accent4
	clrs[7] = *color
	// Accent5
	clrs[8] = *color

	color, _ = NewColor()
	color.Set_Color_B(0)
	color.Set_Color_R(255)
	color.Set_Color_G(255)
	// Accent6
	clrs[9] = *color
	// Hyperlink
	clrs[10] = *color
	// Followed Hyperlink
	clrs[11] = *color

	// Apply custom theme colors on workbook
	wb.CustomTheme("AnyTheme", clrs)

	// Save the workbook
	wb.Save_String(outputApplyCustomThemeColorsOfWorkbookUsingArrayOfColors)

	// Show successful execution message on console
	fmt.Println("ApplyCustomThemeColorsOfWorkbookUsingArrayOfColors executed successfully.\r\n\r\n")
}
