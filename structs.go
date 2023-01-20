package main

//JSON STRUCTS TO GET THE STORE NUMBERS//

type storeNumber struct {
	Name  string `json:"name" bson:"name"`
	Value int    `json:"value" bson:"value"`
	Sales int    `json:"sales" bson:"sales"`
}

type storeNumbers struct {
	Store      string        `json:"store" bson:"store"`
	Month      string        `json:"month" bson:"month"`
	Year       string        `json:"year" bson:"year"`
	Parents    []storeNumber `json:"parents" bson:"parents"`
	Brands     []storeNumber `json:"brands" bson:"brands"`
	Variations []storeNumber `json:"variations" bson:"variations"`
}

// STRUCTS TO KEEP TAB OF PRODUCT EXCEL ROW //

type tablePosition struct {
	// Name     string `json:"name" bson:"name"`
	Position int `json:"position" bson:"position"`
	Format   int `json:"format" bson:"format"`
}

// EXCEL SKELETON //

type monthColumns struct {
	Listings   string
	Sales      string
	Percentage string
	Conversion string
	Separator  string
}
