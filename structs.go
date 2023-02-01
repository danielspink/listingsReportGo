package main

//JSON STRUCTS TO GET THE STORE NUMBERS//

type storeNumber struct {
	Name       string  `json:"name" bson:"name"`
	Listings   int     `json:"value" bson:"value"`
	Sales      int     `json:"sales" bson:"sales"`
	Percentage float64 `json:"percentage" bson:"percentage"`
	Conversion float64 `json:"conversion" bson:"conversion"`
}

type storeNumbers struct {
	Store           string        `json:"store" bson:"store"`
	Month           int           `json:"month" bson:"month"`
	MonthName       string        `json:"monthname" bson:"monthname"`
	Year            int           `json:"year" bson:"year"`
	TotalSales      int           `json:"totalsales" bson:"totalsales"`
	TotalParents    int           `json:"totalparents" bson:"totalparents"`
	TotalBrands     int           `json:"totalbrands" bson:"totalbrands"`
	TotalVariations int           `json:"totalvariations" bson:"totalvariations"`
	SalesPercentage float64       `json:"salespercentage" bson:"salespercentage"`
	SalesConversion float64       `json:"salesconversion" bson:"salesconversion"`
	Parents         []storeNumber `json:"parents" bson:"parents"`
	Brands          []storeNumber `json:"brands" bson:"brands"`
	Variations      []storeNumber `json:"variations" bson:"variations"`
}

// STRUCTS TO KEEP TAB OF PRODUCT EXCEL ROW //

type tablePosition struct {
	Name      string `json:"name" bson:"name"`
	Position  int    `json:"position" bson:"position"`
	Format    int    `json:"format" bson:"format"`
	Separator int    `json:"separator" bson:"separator"`
}

// EXCEL SKELETON //

type monthColumns struct {
	Listings   string
	Sales      string
	Percentage string
	Conversion string
	Separator  string
}
