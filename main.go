package main

func main() {
	setup()
	walmart := storeNumbersFromJson("C:\\Users\\danie\\go\\src\\listings_report\\walmart_Data_december_2022.json")
	// fmt.Println(walmart)
	// runCom()
	buildReportbase(walmart)
}
