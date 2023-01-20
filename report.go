package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func buildReportbase(stores []storeNumbers) {
	filename := "Listings_" + year + ".xlsx"
	xlsx := excelize.NewFile()

	makeSheets(xlsx, stores)
	indexes := makeRowIndxs(xlsx, stores)
	// fmt.Print(indexes)

	for _, data := range stores {
		storeSheet := strings.Title(data.Store)

		makeDataTables(xlsx, storeSheet, data, indexes)
		// 	insertDataToTablesByMonth(xlsx, storeSheet, data.Parents, indexes)
		// 	insertDataToTablesByMonth(xlsx, storeSheet, data.Brands, indexes)
		// 	insertDataToTablesByMonth(xlsx, storeSheet, data.Variations, indexes)
	}

	if err := xlsx.SaveAs(filename); err != nil {
		fmt.Println(err)
	}
}

func makeDataTables(xlsx *excelize.File, sheet string, data storeNumbers, indexes map[string]map[string]tablePosition) {
	//First Column of the sheet with the product Keynames
	currentStore := strings.Title(data.Store)
	// BLUE TABLE //
	rowIndx := 1
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "C"+strconv.Itoa(rowIndx), format(xlsx, "mainTitleCenter"), strings.ToUpper(data.Store))
	rowIndx = rowIndx + 2
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "normalTextLeft"), "Listings")
	rowIndx++
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "blueTextTop"), "")
	rowIndx++
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "blueTextBottom"), data.Store)

	rowIndx = rowIndx + 2

	// PURPLE TABLE //
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "normalTextLeft"), "Products")
	rowIndx++
	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "purpleTextTop"), "")
	rowIndx++
	for _, item := range data.Parents {
		currentRow := indexes[currentStore][item.Name].Position
		currentformat := indexes[currentStore][item.Name].Format
		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(currentRow), "A"+strconv.Itoa(currentRow), currentformat, item.Name)
		// switch numType(indx) {
		// case "Even":
		// 	currentformat := format(xlsx, "purpleTextMid")
		// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
		// 	indexes[item.Name] = tablePosition{rowIndx, currentformat}
		// 	rowIndx++
		// case "Odd":
		// 	currentformat := format(xlsx, "normalTextLeft")
		// 	insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
		// 	indexes[item.Name] = tablePosition{rowIndx, currentformat}
		// 	rowIndx++
		// }
	}
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "purpleTextBottom"), "All")
	// rowIndx = rowIndx + 2

	// GREEN TABLE //
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "normalTextLeft"), "Brands")
	// rowIndx++
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "greenTextTop"), "")
	// rowIndx++
	// for indx, item := range data.Brands {
	// 	switch numType(indx) {
	// 	case "Even":
	// 		currentformat := format(xlsx, "greenTextMid")
	// 		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
	// 		indexes[item.Name] = tablePosition{rowIndx, currentformat}
	// 		rowIndx++
	// 	case "Odd":
	// 		currentformat := format(xlsx, "normalTextLeft")
	// 		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
	// 		indexes[item.Name] = tablePosition{rowIndx, currentformat}
	// 		rowIndx++
	// 	}
	// }
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "greenTextBottom"), "All")
	// rowIndx = rowIndx + 2

	// // ORANGE TABLE //
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "normalTextLeft"), "Variations")
	// rowIndx++
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "orangeTextTop"), "")
	// rowIndx++
	// for indx, item := range data.Variations {
	// 	switch numType(indx) {
	// 	case "Even":
	// 		currentformat := format(xlsx, "orangeTextMid")
	// 		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
	// 		indexes[item.Name] = tablePosition{rowIndx, currentformat}
	// 		rowIndx++
	// 	case "Odd":
	// 		currentformat := format(xlsx, "normalTextLeft")
	// 		insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), currentformat, item.Name)
	// 		indexes[item.Name] = tablePosition{rowIndx, currentformat}
	// 		rowIndx++
	// 	}
	// }
	// insertDataToExcel(xlsx, sheet, "A"+strconv.Itoa(rowIndx), "A"+strconv.Itoa(rowIndx), format(xlsx, "orangeTextBottom"), "All")
}

func insertDataToTablesByMonth(xlsx *excelize.File, sheet string, values []storeNumber, indexes map[string]tablePosition) {
	for _, item := range values {
		cellColumn := reportSkeleton["current"].Listings
		cellRow := indexes[item.Name].Position
		cellFormat := indexes[item.Name].Format
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, item.Value)

		cellColumn = reportSkeleton["current"].Sales
		insertDataToExcel(xlsx, sheet, cellColumn+strconv.Itoa(cellRow), cellColumn+strconv.Itoa(cellRow), cellFormat, item.Sales)
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
		"font":{"color":"#ffffff","size":18,"bold":false},
		"fill":{"type":"pattern","color":["#4A86E8"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"center","wrap_text":false}
	}`)
	formats["mainTitleCenter"] = mainTitleCenter

	blueTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#5b95f9"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["blueTextTop"] = blueTextTop

	blueTextBottom, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#acc9fe"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":false,"text_rotation":0,"horizontal":"left","wrap_text":false}
	}`)
	formats["blueTextBottom"] = blueTextBottom

	purpleTextTop, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":false},
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
		"font":{"color":"#FFFFFF","size":11,"bold":false},
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
		"font":{"color":"#FFFFFF","size":11,"bold":false},
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
