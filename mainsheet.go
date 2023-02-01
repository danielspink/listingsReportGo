package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var uStores map[string]int
var uParents map[string]int
var uBrands map[string]int
var uVars map[string]int

func mainSheetIndexes(data []storeNumbers) map[string]tablePosition {
	// Get the Unique values accross ALL stores to make the main Sheet //
	mainIndexes := make(map[string]tablePosition)
	lStores := make(map[string]int)
	lParents := make(map[string]int)
	lBrands := make(map[string]int)
	lVars := make(map[string]int)

	// rowIndx := 3
	for _, store := range data {
		if store.MonthName != "current" {
			if _, storeExists := lStores[store.Store]; storeExists {
				lStores[store.Store] += store.TotalSales
			} else {
				lStores[store.Store] = store.TotalSales
			}
			for _, parent := range store.Parents {
				if _, parentExists := lParents[parent.Name]; parentExists {
					lParents[parent.Name] += parent.Sales
				} else {
					lParents[parent.Name] = parent.Sales
				}
			}
			for _, brand := range store.Brands {
				if _, brandExists := lBrands[brand.Name]; brandExists {
					lBrands[brand.Name] += brand.Sales
				} else {
					lBrands[brand.Name] = brand.Sales
				}
			}
			for _, variation := range store.Variations {
				if _, varExists := lVars[variation.Name]; varExists {
					lVars[variation.Name] += variation.Sales
				} else {
					lVars[variation.Name] = variation.Sales
				}
			}
		}
	}

	mainIndexes["mainTitle"] = tablePosition{"Listings", 4, Formats["normalTextLeft"], Formats["normalTextLeft"]}
	mainIndexes["mainHeader"] = tablePosition{"", 5, Formats["blueTextTop"], Formats["blueTextTop"]}
	rowIndx := 6

	for _, store := range sortMapValues(lStores) {
		switch numType(rowIndx) {
		case "Even":
			currentformat := Formats["blueTextMid"]
			mainIndexes[store] = tablePosition{store, rowIndx, currentformat, Formats["blueTextTop"]}
			rowIndx++
		case "Odd":
			currentformat := Formats["normalTextLeft"]
			mainIndexes[store] = tablePosition{store, rowIndx, currentformat, Formats["blueTextTop"]}
			rowIndx++
		}
	}

	mainIndexes["mainBottom"] = tablePosition{"All", rowIndx, Formats["blueTextBottom"], Formats["blueTextBottom"]}
	rowIndx = rowIndx + 2
	mainIndexes["parentTitle"] = tablePosition{"Parents", rowIndx, Formats["normalTextLeft"], Formats["normalTextLeft"]}
	rowIndx++
	mainIndexes["parentHeader"] = tablePosition{"", rowIndx, Formats["purpleTextTop"], Formats["purpleTextTop"]}
	rowIndx++
	for _, parent := range sortMapValues(lParents) {
		switch numType(rowIndx) {
		case "Even":
			currentformat := Formats["purpleTextMid"]
			mainIndexes[parent] = tablePosition{parent, rowIndx, currentformat, Formats["purpleTextTop"]}
			rowIndx++
		case "Odd":
			currentformat := Formats["normalTextLeft"]
			mainIndexes[parent] = tablePosition{parent, rowIndx, currentformat, Formats["purpleTextTop"]}
			rowIndx++
		}
	}
	mainIndexes["parentBottom"] = tablePosition{"All", rowIndx, Formats["purpleTextBottom"], Formats["purpleTextBottom"]}
	rowIndx = rowIndx + 2
	mainIndexes["brandTitle"] = tablePosition{"Brands", rowIndx, Formats["normalTextLeft"], Formats["normalTextLeft"]}
	rowIndx++
	mainIndexes["brandHeader"] = tablePosition{"", rowIndx, Formats["greenTextTop"], Formats["greenTextTop"]}
	rowIndx++
	for _, brand := range sortMapValues(lBrands) {
		switch numType(rowIndx) {
		case "Even":
			currentformat := Formats["greenTextMid"]
			mainIndexes[brand] = tablePosition{brand, rowIndx, currentformat, Formats["greenTextTop"]}
			rowIndx++
		case "Odd":
			currentformat := Formats["normalTextLeft"]
			mainIndexes[brand] = tablePosition{brand, rowIndx, currentformat, Formats["greenTextTop"]}
			rowIndx++
		}
	}
	mainIndexes["brandBottom"] = tablePosition{"All", rowIndx, Formats["greenTextBottom"], Formats["greenTextBottom"]}
	rowIndx = rowIndx + 2
	mainIndexes["variationTitle"] = tablePosition{"Variations", rowIndx, Formats["normalTextLeft"], Formats["normalTextLeft"]}
	rowIndx++
	mainIndexes["variationHeader"] = tablePosition{"", rowIndx, Formats["orangeTextTop"], Formats["orangeTextTop"]}
	rowIndx++
	for _, variation := range sortMapValues(lVars) {
		switch numType(rowIndx) {
		case "Even":
			currentformat := Formats["orangeTextMid"]
			mainIndexes[variation] = tablePosition{variation, rowIndx, currentformat, Formats["orangeTextTop"]}
			rowIndx++
		case "Odd":
			currentformat := Formats["normalTextLeft"]
			mainIndexes[variation] = tablePosition{variation, rowIndx, currentformat, Formats["orangeTextTop"]}
			rowIndx++
		}
	}
	mainIndexes["variationBottom"] = tablePosition{"All", rowIndx, Formats["orangeTextBottom"], Formats["orangeTextBottom"]}
	uStores = lStores
	uParents = lParents
	uBrands = lBrands
	uVars = lVars
	return mainIndexes
}

func makeMainSheet(xlsx *excelize.File, data []storeNumbers, mainIndexes map[string]tablePosition) {

	maxListings := make(map[string]int)
	maxParents := make(map[string]int)
	maxBrands := make(map[string]int)
	maxVars := make(map[string]int)
	allSales := make(map[string]int)

	sheet := "Main"
	xlsx.SetColWidth(sheet, "A", "A", 25)
	insertDataToExcel(xlsx, sheet, "B2", "T2", Formats["mainTitleCenter"], "")
	for _, tablePos := range mainIndexes {
		position := "A" + strconv.Itoa(tablePos.Position)
		insertDataToExcel(xlsx, sheet, position, position, tablePos.Format, tablePos.Name)
	}
	for _, store := range data {
		if _, exists := maxListings[store.MonthName]; exists {
			if maxListings[store.MonthName] < store.TotalBrands {
				maxListings[store.MonthName] = store.TotalBrands
			}
		} else {
			maxListings[store.MonthName] = store.TotalBrands
		}

		if _, exists := allSales[store.MonthName]; exists {
			allSales[store.MonthName] += store.TotalSales
		} else {
			allSales[store.MonthName] = store.TotalSales
		}

		for name := range uStores {
			cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "")
			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "")
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "")
			cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "")
			cellPosition = reportSkeleton[store.MonthName].Separator + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes["mainHeader"].Format, "")
		}
		for name := range uParents {
			cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Separator + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes["parentHeader"].Format, "")
		}
		for name := range uBrands {
			cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Separator + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes["brandHeader"].Format, "")
		}
		for name := range uVars {
			cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, 0)
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, "0.00%")
			cellPosition = reportSkeleton[store.MonthName].Separator + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes["variationHeader"].Format, "")
		}
	}
	for _, store := range data {

		insertHeaders(xlsx, sheet, store, mainIndexes)

		cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[store.Store].Position)
		currentListings, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
		if err != nil {
			currentListings = 0
		}
		if store.TotalBrands > currentListings {
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[store.Store].Format, store.TotalBrands)
		}

		cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[store.Store].Position)
		if store.MonthName == "current" {
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[store.Store].Format, uStores[store.Store])
		} else {
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[store.Store].Format, store.TotalSales)
		}

		for _, parent := range store.Parents {
			cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[parent.Name].Position)
			currentListings, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentListings = 0
			}

			if parent.Listings > currentListings {
				if _, exists := maxParents[store.MonthName]; exists {
					maxParents[store.MonthName] -= currentListings
					maxParents[store.MonthName] += parent.Listings
				} else {
					maxParents[store.MonthName] = parent.Listings
				}
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[parent.Name].Format, parent.Listings)
			}

			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[parent.Name].Position)
			currentSales, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentSales = 0
			}
			if store.MonthName != "current" {
				salesInput := currentSales + parent.Sales
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[parent.Name].Format, salesInput)
			} else {
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[parent.Name].Format, uParents[parent.Name])
			}
		}

		for _, brand := range store.Brands {
			cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[brand.Name].Position)
			currentListings, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentListings = 0
			}

			if brand.Listings > currentListings {
				if _, exists := maxBrands[store.MonthName]; exists {
					maxBrands[store.MonthName] -= currentListings
					maxBrands[store.MonthName] += brand.Listings
				} else {
					maxBrands[store.MonthName] = brand.Listings
				}
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[brand.Name].Format, brand.Listings)
			}

			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[brand.Name].Position)
			currentSales, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentSales = 0
			}
			if store.MonthName != "current" {
				salesInput := currentSales + brand.Sales
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[brand.Name].Format, salesInput)
			} else {
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[brand.Name].Format, uBrands[brand.Name])
			}
		}

		for _, variation := range store.Variations {
			cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[variation.Name].Position)
			currentListings, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentListings = 0
			}

			if variation.Listings > currentListings {
				if _, exists := maxVars[store.MonthName]; exists {
					maxVars[store.MonthName] -= currentListings
					maxVars[store.MonthName] += variation.Listings
				} else {
					maxVars[store.MonthName] = variation.Listings
				}
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[variation.Name].Format, variation.Listings)
			}

			cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[variation.Name].Position)
			currentSales, err := strconv.Atoi(xlsx.GetCellValue(sheet, cellPosition))
			if err != nil {
				currentSales = 0
			}
			if store.MonthName != "current" {
				salesInput := currentSales + variation.Sales
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[variation.Name].Format, salesInput)
			} else {
				insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[variation.Name].Format, uVars[variation.Name])
			}
		}
	}

	for _, store := range data {

		indx := mainIndexes["mainBottom"]
		cellPosition := reportSkeleton[store.MonthName].Listings + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, maxListings[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, allSales[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, "")

		cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(indx.Position)
		totalConversion := fmt.Sprintf("%.2f%%", roundFloat((float64(allSales[store.MonthName])*100/float64(maxListings[store.MonthName])), 2))
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, totalConversion)

		/////////////////////////////////
		for name := range uStores {
			listingsCell := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			listings, err := strconv.Atoi(xlsx.GetCellValue(sheet, listingsCell))
			if err != nil {
				listings = 0
			}
			salesCell := reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			sales, err := strconv.Atoi(xlsx.GetCellValue(sheet, salesCell))
			if err != nil {
				listings = 0
			}
			conversionRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(listings)), 2))
			cellPosition := reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, conversionRate)

			percentageRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(allSales[store.MonthName])), 2))
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, percentageRate)
		}

		indx = mainIndexes["parentBottom"]
		cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, maxParents[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, allSales[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, "")

		cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(indx.Position)
		totalConversion = fmt.Sprintf("%.2f%%", roundFloat((float64(allSales[store.MonthName])*100/float64(maxParents[store.MonthName])), 2))
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, totalConversion)

		for name := range uParents {
			listingsCell := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			listings, err := strconv.Atoi(xlsx.GetCellValue(sheet, listingsCell))
			if err != nil {
				listings = 0
			}
			salesCell := reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			sales, err := strconv.Atoi(xlsx.GetCellValue(sheet, salesCell))
			if err != nil {
				listings = 0
			}
			conversionRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(listings)), 2))
			cellPosition := reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, conversionRate)

			percentageRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(allSales[store.MonthName])), 2))
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, percentageRate)
		}

		indx = mainIndexes["brandBottom"]
		cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, maxBrands[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, allSales[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, "")

		cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(indx.Position)
		totalConversion = fmt.Sprintf("%.2f%%", roundFloat((float64(allSales[store.MonthName])*100/float64(maxBrands[store.MonthName])), 2))
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, totalConversion)

		for name := range uBrands {
			listingsCell := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			listings, err := strconv.Atoi(xlsx.GetCellValue(sheet, listingsCell))
			if err != nil {
				listings = 0
			}
			salesCell := reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			sales, err := strconv.Atoi(xlsx.GetCellValue(sheet, salesCell))
			if err != nil {
				listings = 0
			}
			conversionRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(listings)), 2))
			cellPosition := reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, conversionRate)

			percentageRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(allSales[store.MonthName])), 2))
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, percentageRate)
		}

		indx = mainIndexes["variationBottom"]
		cellPosition = reportSkeleton[store.MonthName].Listings + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, maxVars[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Sales + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, allSales[store.MonthName])

		cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(indx.Position)
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, "")

		cellPosition = reportSkeleton[store.MonthName].Conversion + strconv.Itoa(indx.Position)
		totalConversion = fmt.Sprintf("%.2f%%", roundFloat((float64(allSales[store.MonthName])*100/float64(maxVars[store.MonthName])), 2))
		insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, indx.Format, totalConversion)

		for name := range uVars {
			listingsCell := reportSkeleton[store.MonthName].Listings + strconv.Itoa(mainIndexes[name].Position)
			listings, err := strconv.Atoi(xlsx.GetCellValue(sheet, listingsCell))
			if err != nil {
				listings = 0
			}
			salesCell := reportSkeleton[store.MonthName].Sales + strconv.Itoa(mainIndexes[name].Position)
			sales, err := strconv.Atoi(xlsx.GetCellValue(sheet, salesCell))
			if err != nil {
				listings = 0
			}
			conversionRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(listings)), 2))
			cellPosition := reportSkeleton[store.MonthName].Conversion + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, conversionRate)

			percentageRate := fmt.Sprintf("%.2f%%", roundFloat((float64(sales)*100/float64(allSales[store.MonthName])), 2))
			cellPosition = reportSkeleton[store.MonthName].Percentage + strconv.Itoa(mainIndexes[name].Position)
			insertDataToExcel(xlsx, sheet, cellPosition, cellPosition, mainIndexes[name].Format, percentageRate)
		}
	}
}
