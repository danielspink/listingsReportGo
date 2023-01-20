package main

func main() {
	setup()
	stores := getStoreDataByYear(year)
	buildReportbase(stores)
}
