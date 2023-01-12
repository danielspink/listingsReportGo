package main

//JSON STRUCTS TO GET THE STORE NUMBERS//

type storeNumber struct {
	Name  string `json:"name"`
	Value int    `json:"value"`
}

type storeNumbers struct {
	Parents    []storeNumber `json:"parents"`
	Brands     []storeNumber `json:"brands"`
	Variations []storeNumber `json:"variations"`
}
