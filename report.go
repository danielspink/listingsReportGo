package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func buildReportbase(data storeNumbers) {
	filename := "Listings_" + year + "_" + month + ".xlsx"
	xlsx := excelize.NewFile()

	storeSheet := "Walmart"
	xlsx.NewSheet(storeSheet)

	cellindx := 2
	textformat := format(xlsx, "purpleTextCenter")
	//build parents//
	for _, arr := range data.Parents {
		insertDataToExcel(xlsx, storeSheet, "B"+strconv.Itoa(cellindx), "B"+strconv.Itoa(cellindx), textformat, arr.Name)
		insertDataToExcel(xlsx, storeSheet, "C"+strconv.Itoa(cellindx), "C"+strconv.Itoa(cellindx), textformat, strconv.Itoa(arr.Value))
		cellindx++
	}
	if err := xlsx.SaveAs(filename); err != nil {
		fmt.Println(err)
	}
}

func insertDataToExcel(xlsx *excelize.File, sheet string, firstCol string, lastCol string, format int, text string) {
	xlsx.SetCellValue(sheet, firstCol, text)
	xlsx.MergeCell(sheet, firstCol, lastCol)
	xlsx.SetCellStyle(sheet, firstCol, lastCol, format)
}

func format(xlsx *excelize.File, format string) int {
	formats := map[string]int{}
	colorTheme := "#0560E2"
	alternatedColorTheme := "#34507f"

	title, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":24,"bold":true},
		"fill":{"type":"pattern","color":["` + colorTheme + `"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"","wrap_text":true}
	}`)
	formats["title"] = title

	titleCentered, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":24,"bold":true},
		"fill":{"type":"pattern","color":["` + colorTheme + `"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["titleCentered"] = titleCentered

	subData, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":18,"bold":true},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"left","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"","wrap_text":true}
	}`)
	formats["subData"] = subData

	quarter, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":18,"bold":true},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["quarter"] = quarter

	subDataLeft, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":14,"bold":true},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"left","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"","wrap_text":true}
	}`)
	formats["subDataLeft"] = subDataLeft

	subDataCenter, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":14,"bold":true},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"left","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["subDataCenter"] = subDataCenter

	subDataRight, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":14,"bold":true},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"right","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"right","wrap_text":true}
	}`)
	formats["subDataRight"] = subDataRight

	subTitle, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["` + colorTheme + `"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["subTitle"] = subTitle

	subTitleAlternated, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["` + alternatedColorTheme + `"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"center","wrap_text":true}
	}`)
	formats["subTitleAlternated"] = subTitleAlternated

	subTitleLeft, _ := xlsx.NewStyle(`{
		"font":{"color":"#FFFFFF","size":11,"bold":true},
		"fill":{"type":"pattern","color":["` + colorTheme + `"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"left","wrap_text":true}
	}`)
	formats["subTitleLeft"] = subTitleLeft

	normalTextRight, _ := xlsx.NewStyle(`{
		"font":{"color":"#000000","size":11,"bold":false},
		"fill":{"type":"pattern","color":["#ffffff"],"pattern":1},
		"alignment":{"vertical":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":0,"horizontal":"right","wrap_text":true}
	}`)
	formats["normalTextRight"] = normalTextRight

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

	return formats[format]
}
