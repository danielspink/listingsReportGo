package main

import (
	"fmt"
	"sort"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func currentStoreTable(store []storeNumbers) {

}

func makeRowIndxs(xlsx *excelize.File, data []storeNumbers) map[string]map[string]tablePosition {
	indexes := make(map[string]map[string]tablePosition)
	storeproducts := make(map[string]map[string]map[string]int)
	for _, store := range data {
		currentSheet := strings.Title(store.Store)
		if _, storeExists := storeproducts[currentSheet]; storeExists {
		} else {
			storeproducts[currentSheet] = make(map[string]map[string]int)
		}
		if _, parentsExist := storeproducts[currentSheet]["Parents"]; parentsExist {
		} else {
			storeproducts[currentSheet]["Parents"] = make(map[string]int)
		}
		for _, parent := range store.Parents {
			if _, parentNameExists := storeproducts[currentSheet]["Parents"]; parentNameExists {
				storeproducts[currentSheet]["Parents"][parent.Name] += parent.Sales
			} else {
				storeproducts[currentSheet]["Parents"][parent.Name] = parent.Sales
			}
		}
		sortMapValues(storeproducts[currentSheet]["Parents"])
	}

	for _, store := range data {
		rowIndx := 9
		currentSheet := strings.Title(store.Store)
		if _, storeOk := indexes[currentSheet]; storeOk {
			for _, parent := range store.Parents {

				if _, parentOk := indexes[currentSheet][parent.Name]; parentOk {
					continue
				} else {
					switch numType(rowIndx) {
					case "Even":
						currentformat := format(xlsx, "purpleTextMid")
						indexes[currentSheet][parent.Name] = tablePosition{rowIndx, currentformat}
						// fmt.Printf("%v %v \n", currentSheet, len(indexes[currentSheet]))
						rowIndx++
					case "Odd":
						currentformat := format(xlsx, "normalTextLeft")
						indexes[currentSheet][parent.Name] = tablePosition{rowIndx, currentformat}
						// fmt.Printf("%v %v \n", currentSheet, len(indexes[currentSheet]))
						rowIndx++
					}
				}
			}
		} else {

			indexes[currentSheet] = make(map[string]tablePosition)

			for indx, parent := range store.Parents {

				if _, parentOk := indexes[currentSheet][parent.Name]; parentOk {
					continue
				} else {
					switch numType(indx) {
					case "Even":
						currentformat := format(xlsx, "purpleTextMid")
						indexes[currentSheet][parent.Name] = tablePosition{rowIndx, currentformat}
						// fmt.Printf("%v %v \n", currentSheet, len(indexes[currentSheet]))
						rowIndx++
					case "Odd":
						currentformat := format(xlsx, "normalTextLeft")
						indexes[currentSheet][parent.Name] = tablePosition{rowIndx, currentformat}
						// fmt.Printf("%v %v \n", currentSheet, len(indexes[currentSheet]))
						rowIndx++
					}
				}
			}
		}
	}
	return indexes
}

func sortMapValues(inMap map[string]int) map[string]int {
	outMap := make(map[string]int)
	keys := make([]string, 0, len(inMap))
	for key := range inMap {
		keys = append(keys, key)
	}
	sort.SliceStable(keys, func(i, j int) bool {
		return inMap[keys[i]] > inMap[keys[j]]
	})
	for _, k := range keys {
		outMap[k] = inMap[k]
	}

	for k, v := range outMap {
		fmt.Println(k, v)
	}
	return outMap
}
