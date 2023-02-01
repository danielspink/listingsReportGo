package main

import (
	"context"
	"fmt"
	"os"
	"strconv"

	"github.com/SmartPrintsInk/crashdis"
	"github.com/SmartPrintsInk/spingo"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/mongo"
)

func getStoreDataByMonth(month string, year int) (stores []storeNumbers) {
	client, err := spingo.AccessFor(os.Getenv("HostIP"))
	crashdis.CrashDis(err, "mongodb connection")
	defer spingo.Close()
	collection := client.Database("reports").Collection("listingsByStore")

	matchStage := bson.D{
		{Key: "$match", Value: bson.D{
			{Key: "month", Value: month},
			{Key: "year", Value: year},
		}},
	}
	projectStage := bson.D{{Key: "$project", Value: bson.D{
		{Key: "_id", Value: 0},
	}}}

	var results []storeNumbers

	pipeline := mongo.Pipeline{
		matchStage,
		projectStage,
	}

	cursor, err := collection.Aggregate(context.TODO(), pipeline)

	if err == mongo.ErrNoDocuments {
		fmt.Printf("No documents were found in %s %s\n", month, strconv.Itoa(year))
		return
	}
	crashdis.CheckDis(err, "Mongo Document Search")
	if err = cursor.All(context.TODO(), &results); err != nil {
		crashdis.CrashDis(err, "Getting documents")
	}
	return results
}

func getStoreDataByYear(year int) (stores []storeNumbers) {
	client, err := spingo.AccessFor(os.Getenv("HostIP"))
	crashdis.CrashDis(err, "mongodb connection")
	defer spingo.Close()
	collection := client.Database("reports").Collection("listingsByStore")

	matchStage := bson.D{
		{Key: "$match", Value: bson.D{
			{Key: "year", Value: year},
		}},
	}
	projectStage := bson.D{{Key: "$project", Value: bson.D{
		{Key: "_id", Value: 1},
		{Key: "store", Value: 1},
		{Key: "month", Value: 1},
		{Key: "monthname", Value: 1},
		{Key: "year", Value: 1},
		{Key: "totalsales", Value: bson.D{{Key: "$sum", Value: "$parents.sales"}}},
		{Key: "totalparents", Value: bson.D{{Key: "$sum", Value: "$parents.value"}}},
		{Key: "totalbrands", Value: bson.D{{Key: "$sum", Value: "$brands.value"}}},
		{Key: "totalvariations", Value: bson.D{{Key: "$sum", Value: "$variations.value"}}},
		{Key: "parents", Value: 1},
		{Key: "brands", Value: 1},
		{Key: "variations", Value: 1},
	}}}

	var results []storeNumbers

	pipeline := mongo.Pipeline{
		matchStage,
		projectStage,
	}

	cursor, err := collection.Aggregate(context.TODO(), pipeline)

	if err == mongo.ErrNoDocuments {
		fmt.Printf("No documents were found in %s %s\n", month, strconv.Itoa(year))
		return
	}
	crashdis.CheckDis(err, "Mongo Document Search")
	if err = cursor.All(context.TODO(), &results); err != nil {
		crashdis.CrashDis(err, "Getting documents")
	}

	if len(results) > 0 {
		results = percentajeRatesByMonth(makeCurrentData(results))
		return results
	}
	return
}

func makeCurrentData(data []storeNumbers) []storeNumbers {
	// Armar current tables para el excel //
	currents := make(map[string]storeNumbers)
	parents := make(map[string]map[string]storeNumber)
	brands := make(map[string]map[string]storeNumber)
	variations := make(map[string]map[string]storeNumber)

	allParents := make(map[string]map[string]storeNumber)
	allBrands := make(map[string]map[string]storeNumber)
	allVars := make(map[string]map[string]storeNumber)

	for _, store := range data {
		if _, exists := allParents[store.Store]; exists {
		} else {
			allParents[store.Store] = make(map[string]storeNumber)
			allBrands[store.Store] = make(map[string]storeNumber)
			allVars[store.Store] = make(map[string]storeNumber)
		}

		for _, item := range store.Parents {
			allParents[store.Store][item.Name] = storeNumber{Name: item.Name}
		}
		for _, item := range store.Brands {
			allBrands[store.Store][item.Name] = storeNumber{Name: item.Name}
		}
		for _, item := range store.Variations {
			allVars[store.Store][item.Name] = storeNumber{Name: item.Name}
		}
	}

	for indx, store := range data {
		tempParents := make(map[string]int)
		for _, parent := range store.Parents {
			tempParents[parent.Name] = 1
		}
		for name, item := range allParents[store.Store] {
			if _, exists := tempParents[name]; exists {
			} else {
				store.Parents = append(store.Parents, item)
			}
		}

		tempBrands := make(map[string]int)
		for _, brand := range store.Brands {
			tempBrands[brand.Name] = 1
		}
		for name, item := range allBrands[store.Store] {
			if _, exists := tempBrands[name]; exists {
			} else {
				store.Brands = append(store.Brands, item)
			}
		}

		tempVars := make(map[string]int)
		for _, vari := range store.Variations {
			tempVars[vari.Name] = 1
		}
		for name, item := range allVars[store.Store] {
			if _, exists := tempVars[name]; exists {
			} else {
				store.Variations = append(store.Variations, item)
			}
		}

		data[indx].Parents = store.Parents
		data[indx].Brands = store.Brands
		data[indx].Variations = store.Variations

		if entry, exists := currents[store.Store]; exists {
			if store.Month > entry.Month {
				entry.TotalParents = store.TotalParents
				entry.TotalBrands = store.TotalBrands
				entry.TotalVariations = store.TotalVariations
			}
			entry.TotalSales += store.TotalSales
			currents[store.Store] = entry

		} else {
			currents[store.Store] = store
			parents[store.Store] = make(map[string]storeNumber)
			brands[store.Store] = make(map[string]storeNumber)
			variations[store.Store] = make(map[string]storeNumber)
		}

		newParents := make(map[string]int)
		for _, parent := range store.Parents {
			newParents[parent.Name] = parent.Listings
			if entry, exists := parents[store.Store][parent.Name]; exists {
				if store.Month > currents[store.Store].Month {
					entry.Listings = parent.Listings
				}
				entry.Sales = entry.Sales + parent.Sales
				parents[store.Store][parent.Name] = entry
			} else {
				parents[store.Store][parent.Name] = parent
			}
		}
		for name, entry := range parents[store.Store] {
			if _, parentExists := newParents[name]; parentExists {
			} else {
				entry.Listings = 0
				parents[store.Store][name] = entry
			}

		}

		newBrands := make(map[string]int)
		for _, brand := range store.Brands {
			newBrands[brand.Name] = brand.Listings
			if entry, exists := brands[store.Store][brand.Name]; exists {
				if store.Month > currents[store.Store].Month {
					entry.Listings = brand.Listings
				}
				entry.Sales = entry.Sales + brand.Sales
				brands[store.Store][brand.Name] = entry
			} else {
				brands[store.Store][brand.Name] = brand
			}
		}
		for name, entry := range brands[store.Store] {
			if _, brandExists := newBrands[name]; brandExists {
			} else {
				entry.Listings = 0
				brands[store.Store][name] = entry
			}
		}

		newVars := make(map[string]int)
		for _, variation := range store.Variations {
			newVars[variation.Name] = variation.Listings
			if varEntry, varExists := variations[store.Store][variation.Name]; varExists {
				if store.Month > currents[store.Store].Month {
					varEntry.Listings = variation.Listings
				}
				varEntry.Sales = varEntry.Sales + variation.Sales
				variations[store.Store][variation.Name] = varEntry
			} else {
				variations[store.Store][variation.Name] = variation
			}
		}
		for name, entry := range variations[store.Store] {
			if _, exists := newVars[name]; exists {
			} else {
				entry.Listings = 0
				variations[store.Store][name] = entry
			}
		}
	}

	for name, info := range currents {
		currentStore := storeNumbers{
			Store:           name,
			Month:           0,
			MonthName:       "current",
			Year:            info.Year,
			TotalSales:      info.TotalSales,
			TotalParents:    info.TotalParents,
			TotalBrands:     info.TotalBrands,
			TotalVariations: info.TotalVariations,
		}
		var parentList []storeNumber
		for _, item := range parents[name] {
			parentList = append(parentList, item)
		}
		currentStore.Parents = parentList

		var brandList []storeNumber
		for _, item := range brands[name] {
			brandList = append(brandList, item)
		}
		currentStore.Brands = brandList

		var varList []storeNumber
		for _, item := range variations[name] {
			varList = append(varList, item)
		}
		currentStore.Variations = varList

		data = append(data, currentStore)
	}

	return data
}
