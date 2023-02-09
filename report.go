package main

import (
	"fmt"
	_ "image/png"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func buildReportbase(stores []storeNumbers) {
	filename := "Listings_" + strconv.Itoa(year) + ".xlsx"
	xlsx := excelize.NewFile()

	makeSheets(xlsx, stores)
	makeFormats(xlsx)
	indexes := makeRowIndxs(xlsx, stores)
	mainIndxs := mainSheetIndexes(stores)
	makeMainSheet(xlsx, stores, mainIndxs)

	for _, data := range stores {
		storeSheet := strings.Title(data.Store)

		makeDataTables(xlsx, storeSheet, data, indexes)
		insertDataToTablesByMonth(xlsx, storeSheet, data, indexes[storeSheet])
	}
	xlsx.SetActiveSheet(0)

	if err := xlsx.SaveAs(filename); err != nil {
		fmt.Println(err)
	}
}

func makeDataTables(xlsx *excelize.File, sheet string, data storeNumbers, indexes map[string]map[string]tablePosition) {
	//First Column of the sheet with the product Keynames
	currentStore := strings.Title(data.Store)

	// BLUE TABLE //
	imgLocation := fmt.Sprintf("imgs\\%v.png", currentStore)
	err := xlsx.AddPicture(currentStore, "A1", imgLocation, `{"x_scale":0.5,"y_scale":0.5}`)
	if err != nil {
		fmt.Println(err)
	}

	for _, val := range indexes[currentStore] {
		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(val.Position), "A"+strconv.Itoa(val.Position), val.Format, val.Name)
	}
}

func insertDataToTablesByMonth(xlsx *excelize.File, sheet string, data storeNumbers, indexes map[string]tablePosition) {
	insertHeaders(xlsx, sheet, data, indexes)
	processStoreData(xlsx, sheet, data.MonthName, data.Parents, indexes)
	processStoreData(xlsx, sheet, data.MonthName, data.Brands, indexes)
	processStoreData(xlsx, sheet, data.MonthName, data.Variations, indexes)
}

func insertHeaders(xlsx *excelize.File, sheet string, data storeNumbers, indexes map[string]tablePosition) {
	for name, value := range indexes {
		if strings.Contains(name, "Header") {
			cellColumn := reportSkeleton[data.MonthName].Listings
			cellRow := value.Position
			cellFormat := value.Format
			xlsx.SetColWidth(sheet, cellColumn, cellColumn, float64(len(data.MonthName))+3)
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.MonthName)

			cellColumn = reportSkeleton[data.MonthName].Sales
			xlsx.SetColWidth(sheet, cellColumn, cellColumn, float64(len(data.MonthName+" Sales"))+3)
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.MonthName+" Sales")

			cellColumn = reportSkeleton[data.MonthName].Percentage
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "%")

			cellColumn = reportSkeleton[data.MonthName].Conversion
			xlsx.SetColWidth(sheet, cellColumn, cellColumn, float64(len("Conversion"))+3)
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "Conversion")

			cellColumn = reportSkeleton[data.MonthName].Separator
			xlsx.SetColWidth(sheet, cellColumn, cellColumn, 2.00)
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "")
		} else if strings.Contains(name, "Bottom") {
			cellRow := value.Position
			cellFormat := value.Format

			cellColumn := reportSkeleton[data.MonthName].Percentage
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "")

			cellColumn = reportSkeleton[data.MonthName].Listings
			xlsx.SetColWidth(sheet, cellColumn, cellColumn, float64(len(data.MonthName))+3)

			switch name {
			case "mainBottom":
				insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.TotalBrands)
				cellColumn = reportSkeleton[data.MonthName].Percentage
				insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, fmt.Sprintf("%.2f%%", data.SalesPercentage))
			case "parentBottom":
				insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.TotalParents)
			case "brandBottom":
				insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.TotalBrands)
			case "variationBottom":
				insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.TotalVariations)
			}

			cellColumn = reportSkeleton[data.MonthName].Sales
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, data.TotalSales)

			cellColumn = reportSkeleton[data.MonthName].Conversion
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, fmt.Sprintf("%.2f%%", data.SalesConversion))
			cellColumn = reportSkeleton[data.MonthName].Separator
			insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "")
		}
	}
}

func processStoreData(xlsx *excelize.File, sheet string, month string, value []storeNumber, indexes map[string]tablePosition) {
	for _, item := range value {
		cellColumn := reportSkeleton[month].Listings
		cellRow := indexes[item.Name].Position
		cellFormat := indexes[item.Name].Format
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, item.Listings)

		cellColumn = reportSkeleton[month].Sales
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, item.Sales)

		cellColumn = reportSkeleton[month].Percentage
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, fmt.Sprintf("%.2f%%", item.Percentage))

		cellColumn = reportSkeleton[month].Conversion
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, fmt.Sprintf("%.2f%%", item.Conversion))

		cellColumn = reportSkeleton[month].Separator
		cellFormat = indexes[item.Name].Separator
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, "")
	}
}

func insertDataToExcel(xlsx *excelize.File, sheet string, firstCol string, lastCol string, format int, text interface{}) {
	xlsx.SetCellValue(sheet, firstCol, text)
	xlsx.MergeCell(sheet, firstCol, lastCol)
	xlsx.SetCellStyle(sheet, firstCol, lastCol, format)
}

func format(xlsx *excelize.File, format string) int {
	formats := map[string]int{}

	normalTextRight, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"right","wrap_text":false}
	}`)
	formats["normalTextRight"] = normalTextRight

	normalTextLeft, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["normalTextLeft"] = normalTextLeft

	normalTextCenter, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["normalTextCenter"] = normalTextCenter

	normalTextRightAlternative, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#e5e5e5"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"right","wrap_text":true}
	}`)
	formats["normalTextRightAlternative"] = normalTextRightAlternative

	normalTextCenterAlternative, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#e5e5e5"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["normalTextCenterAlternative"] = normalTextCenterAlternative

	purpleTextCenter, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#e8e7fc"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["purpleTextCenter"] = purpleTextCenter

	mainTitleCenter, _ := xlsx.NewStyle(`{
		"font":{"color":"#ffffff","size":25,"bold":false},
		"fill":{"type":"pattern","color":["#4A86E8"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"center","wrap_text":false}
	}`)
	formats["mainTitleCenter"] = mainTitleCenter

	mainTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#ffffff","size":20,"bold":true},
		"fill":{"type":"pattern","color":["#5b95f9"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"center","wrap_text":false}
	}`)
	formats["mainTextTop"] = mainTextTop

	blueTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#ffffff","size":11,"bold":true},
		"fill":{"type":"pattern","color":["#5b95f9"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["blueTextTop"] = blueTextTop

	blueTextMid, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#e8f0fe"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["blueTextMid"] = blueTextMid

	blueTextBottom, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#acc9fe"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["blueTextBottom"] = blueTextBottom

	purpleTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["##8989eb"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["purpleTextTop"] = purpleTextTop

	purpleTextMid, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#e8e7fc"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["purpleTextMid"] = purpleTextMid

	purpleTextBottom, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":true},
		"fill":{"type":"pattern","color":["###C4C3F7"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["purpleTextBottom"] = purpleTextBottom

	greenTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["#6aa84f"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["greenTextTop"] = greenTextTop

	greenTextMid, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#eef7e3"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["greenTextMid"] = greenTextMid

	greenTextBottom, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":true},
		"fill":{"type":"pattern","color":["#B6D7A8"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["greenTextBottom"] = greenTextBottom

	orangeTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["#ff9900"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["orangeTextTop"] = orangeTextTop

	orangeTextMid, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#fce5cd"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["orangeTextMid"] = orangeTextMid

	orangeTextBottom, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":true},
		"fill":{"type":"pattern","color":["#FCE8B2"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["orangeTextBottom"] = orangeTextBottom

	return formats[format]
}

func numType(num int) string {
	if num%2 == 0 {
		return "Even"
	} else {
		return "Odd"
	}
}

func makeSheets(xlsx *excelize.File, stores []storeNumbers) {
	xlsx.SetSheetName("Sheet1", "Main")
	for _, data := range stores {
		storeSheet := strings.Title(data.Store)
		xlsx.NewSheet(storeSheet)
		xlsx.SetColWidth(storeSheet, "A", "A", 25)
	}
}
