package main

func main() {
	setup()
	stores := getStoreDataByYear(year)

	// for _, stor := range stores {
	// 	fmt.Printf("%v %v %v\n\n", stor.Store, stor.Month, stor.Year)
	// }
	buildReportbase(stores)
}
