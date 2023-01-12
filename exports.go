package main

import (
	"encoding/json"
	"io/ioutil"
	"log"
	"os"
)

func storeNumbersFromJson(path string) (storeResult storeNumbers) {
	var stor storeNumbers

	jsonF, err := os.Open(path)
	if err != nil {
		log.Fatal(err)
	}
	defer jsonF.Close()

	byteValue, err := ioutil.ReadAll(jsonF)
	if err != nil {
		log.Fatal(err)
	}
	json.Unmarshal(byteValue, &stor)
	storeResult = stor
	return storeResult
}
