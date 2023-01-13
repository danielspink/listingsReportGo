package main

//JSON STRUCTS TO GET THE STORE NUMBERS//

type storeNumber struct {
	Name  string `json:"name" bson:"name"`
	Value int    `json:"value" bson:"value"`
}

type storeNumbers struct {
	Store      string        `json:"store" bson:"store"`
	Month      string        `json:"month" bson:"month"`
	Year       string        `json:"year" bson:"year"`
	Parents    []storeNumber `json:"parents" bson:"parents"`
	Brands     []storeNumber `json:"brands" bson:"brands"`
	Variations []storeNumber `json:"variations" bson:"variations"`
}

type tablePositions struct {
}
