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
		{Key: "monthname", Value: 1},
		{Key: "year", Value: 1},
		{Key: "totalsales", Value: bson.D{{Key: "$sum", Value: "$parents.sales"}}},
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
		return results
	}
	return
}
