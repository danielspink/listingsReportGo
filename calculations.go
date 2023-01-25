package main

import (
	"math"
	"sort"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func currentStoreTable(store []storeNumbers) {
	/// Armar la current table del excel
}

func makeRowIndxs(xlsx *excelize.File, data []storeNumbers) map[string]map[string]tablePosition {
	indexes := make(map[string]map[string]tablePosition)
	storeproducts := make(map[string]map[string]map[string]int)

	// Junta la informacion de la tienda para armar la tabla completa //
	for _, store := range data {
		currentSheet := store.Store
		if _, storeExists := storeproducts[currentSheet]; storeExists {
		} else {
			storeproducts[currentSheet] = make(map[string]map[string]int)
		}
		if _, parentsExist := storeproducts[currentSheet]["Parents"]; parentsExist {
		} else {
			storeproducts[currentSheet]["Parents"] = make(map[string]int)
			storeproducts[currentSheet]["Brands"] = make(map[string]int)
			storeproducts[currentSheet]["Variations"] = make(map[string]int)
		}
		for _, parent := range store.Parents {
			if _, parentNameExists := storeproducts[currentSheet]["Parents"]; parentNameExists {
				storeproducts[currentSheet]["Parents"][parent.Name] += parent.Sales
			} else {
				storeproducts[currentSheet]["Parents"][parent.Name] = parent.Sales
			}
		}
		for _, brand := range store.Brands {
			if _, brandNameExists := storeproducts[currentSheet]["Brands"]; brandNameExists {
				storeproducts[currentSheet]["Brands"][brand.Name] += brand.Sales
			} else {
				storeproducts[currentSheet]["Brands"][brand.Name] = brand.Sales
			}
		}
		for _, variation := range store.Variations {
			if _, brandNameExists := storeproducts[currentSheet]["Variations"]; brandNameExists {
				storeproducts[currentSheet]["Variations"][variation.Name] += variation.Sales
			} else {
				storeproducts[currentSheet]["Variations"][variation.Name] = variation.Sales
			}
		}
	}

	// Asigna el numero de fila y el formato de cada producto y
	for store, item := range storeproducts {
		currentSheet := strings.Title(store)
		if _, storeOk := indexes[currentSheet]; storeOk {
		} else {
			indexes[currentSheet] = make(map[string]tablePosition)
		}
		indexes[currentSheet]["mainTitle"] = tablePosition{"Listings", 3, Formats["normalTextLeft"], Formats["normalTextLeft"]}
		indexes[currentSheet]["mainHeader"] = tablePosition{"", 4, Formats["blueTextTop"], Formats["blueTextTop"]}
		indexes[currentSheet]["mainBottom"] = tablePosition{currentSheet, 5, Formats["blueTextBottom"], Formats["blueTextTop"]}

		rowIndx := 8
		indexes[currentSheet]["parentTitle"] = tablePosition{"Parents", rowIndx - 1, Formats["normalTextLeft"], Formats["normalTextLeft"]}
		rowIndx++
		indexes[currentSheet]["parentHeader"] = tablePosition{"", rowIndx - 1, Formats["purpleTextTop"], Formats["purpleTextTop"]}

		for _, product := range sortMapValues(item["Parents"]) {
			switch numType(rowIndx) {
			case "Even":
				currentformat := Formats["purpleTextMid"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["purpleTextTop"]}
				rowIndx++
			case "Odd":
				currentformat := Formats["normalTextLeft"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["purpleTextTop"]}
				rowIndx++
			}
		}
		indexes[currentSheet]["parentBottom"] = tablePosition{"All", rowIndx, Formats["purpleTextBottom"], Formats["purpleTextTop"]}
		rowIndx = rowIndx + 2
		indexes[currentSheet]["brandTitle"] = tablePosition{"Brands", rowIndx, Formats["normalTextLeft"], Formats["normalTextLeft"]}
		rowIndx = rowIndx + 2
		indexes[currentSheet]["brandHeader"] = tablePosition{"", rowIndx - 1, Formats["greenTextTop"], Formats["greenTextTop"]}

		for _, product := range sortMapValues(item["Brands"]) {
			switch numType(rowIndx) {
			case "Even":
				currentformat := Formats["greenTextMid"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["greenTextTop"]}
				rowIndx++
			case "Odd":
				currentformat := Formats["normalTextLeft"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["greenTextTop"]}
				rowIndx++
			}
		}

		indexes[currentSheet]["brandBottom"] = tablePosition{"All", rowIndx, Formats["greenTextBottom"], Formats["greenTextTop"]}
		rowIndx = rowIndx + 2
		indexes[currentSheet]["variationTitle"] = tablePosition{"Variations", rowIndx, Formats["normalTextLeft"], Formats["normalTextLeft"]}
		rowIndx = rowIndx + 2
		indexes[currentSheet]["variationHeader"] = tablePosition{"", rowIndx - 1, Formats["orangeTextTop"], Formats["orangeTextTop"]}

		for _, product := range sortMapValues(item["Variations"]) {
			switch numType(rowIndx) {
			case "Even":
				currentformat := Formats["orangeTextMid"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["orangeTextTop"]}
				rowIndx++
			case "Odd":
				currentformat := Formats["normalTextLeft"]
				indexes[currentSheet][product] = tablePosition{product, rowIndx, currentformat, Formats["orangeTextTop"]}
				rowIndx++
			}
		}
		indexes[currentSheet]["variationBottom"] = tablePosition{"All", rowIndx, Formats["orangeTextBottom"], Formats["orangeTextTop"]}
	}
	return indexes
}

func sortMapValues(inMap map[string]int) []string {
	keys := make([]string, 0, len(inMap))
	for key := range inMap {
		keys = append(keys, key)
	}
	sort.SliceStable(keys, func(i, j int) bool {
		return inMap[keys[i]] > inMap[keys[j]]
	})
	return keys
}

func makeFormats(xlsx *excelize.File) {
	Formats = map[string]int{
		"blueTextTop":      format(xlsx, "blueTextTop"),
		"blueTextBottom":   format(xlsx, "blueTextBottom"),
		"purpleTextTop":    format(xlsx, "purpleTextTop"),
		"purpleTextMid":    format(xlsx, "purpleTextMid"),
		"purpleTextBottom": format(xlsx, "purpleTextBottom"),
		"greenTextTop":     format(xlsx, "greenTextTop"),
		"greenTextMid":     format(xlsx, "greenTextMid"),
		"greenTextBottom":  format(xlsx, "greenTextBottom"),
		"orangeTextTop":    format(xlsx, "orangeTextTop"),
		"orangeTextMid":    format(xlsx, "orangeTextMid"),
		"orangeTextBottom": format(xlsx, "orangeTextBottom"),
		"normalTextLeft":   format(xlsx, "normalTextLeft"),
	}
}

func percentajeRatesByMonth(data []storeNumbers) []storeNumbers {
	monthlysales := make(map[string]int)

	for storindx, store := range data {
		if _, monthExists := monthlysales[store.Month]; monthExists {
			monthlysales[store.Month] += store.TotalSales
		} else {
			monthlysales[store.Month] = store.TotalSales
		}

		if store.TotalSales > 0 {
			for indx, parent := range store.Parents {
				if parent.Sales > 0 {
					percentageRate := (float64(parent.Sales) * 100) / float64(store.TotalSales)
					data[storindx].Parents[indx].Percentage = roundFloat(percentageRate, 2)
				}
			}
			for indx, brand := range store.Brands {
				if brand.Sales > 0 {
					percentageRate := (float64(brand.Sales) * 100) / float64(store.TotalSales)
					data[storindx].Brands[indx].Percentage = roundFloat(percentageRate, 2)
				}
			}
			for indx, variation := range store.Variations {
				if variation.Sales > 0 {
					percentageRate := (float64(variation.Sales) * 100) / float64(store.TotalSales)
					data[storindx].Variations[indx].Percentage = roundFloat(percentageRate, 2)
				}
			}
		}
	}
	for storindx, store := range data {
		percentageRate := ((float64(store.TotalSales) * 100) / float64(monthlysales[store.Month]))
		data[storindx].SalesPercentage = roundFloat(percentageRate, 2)
	}
	return data
}

func roundFloat(val float64, precision uint) float64 {
	ratio := math.Pow(10, float64(precision))
	return math.Round(val*ratio) / ratio
}
