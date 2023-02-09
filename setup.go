package main

import (
	"flag"
	"strings"
	"time"
)

var reportSkeleton = make(map[string]monthColumns)
var Formats = make(map[string]int)

// var MainSheetInfo
var option string
var store string
var month string
var year int

func setup() {
	loadFlags()
	processDates()
	reportColumns()
}

func loadFlags() {
	flag.StringVar(&option, "o", "none", "options: all/one")
	flag.IntVar(&year, "y", 0, "The Specific year to process")
	flag.StringVar(&month, "m", "none", "The Specific month to process")
	flag.StringVar(&store, "s", "none", "a specific store")
	flag.Parse()
}

func processDates() {
	curMonth := time.Now().Month() + 1
	if month == "none" {
		month = strings.ToLower(curMonth.String())
	}
	if year == 0 {
		year = time.Now().Year()
	}
}

func reportColumns() {
	reportSkeleton["current"] = monthColumns{
		Listings:   "B",
		Sales:      "C",
		Percentage: "D",
		Conversion: "E",
		Separator:  "F",
	}
	reportSkeleton["january"] = monthColumns{
		Listings:   "G",
		Sales:      "H",
		Percentage: "I",
		Conversion: "J",
		Separator:  "K",
	}
	reportSkeleton["february"] = monthColumns{
		Listings:   "L",
		Sales:      "M",
		Percentage: "N",
		Conversion: "O",
		Separator:  "P",
	}
	reportSkeleton["march"] = monthColumns{
		Listings:   "Q",
		Sales:      "R",
		Percentage: "S",
		Conversion: "T",
		Separator:  "U",
	}
	reportSkeleton["april"] = monthColumns{
		Listings:   "V",
		Sales:      "W",
		Percentage: "X",
		Conversion: "Y",
		Separator:  "Z",
	}
	reportSkeleton["may"] = monthColumns{
		Listings:   "AA",
		Sales:      "AB",
		Percentage: "AC",
		Conversion: "AD",
		Separator:  "AE",
	}
	reportSkeleton["june"] = monthColumns{
		Listings:   "AF",
		Sales:      "AG",
		Percentage: "AH",
		Conversion: "AI",
		Separator:  "AJ",
	}
	reportSkeleton["july"] = monthColumns{
		Listings:   "AK",
		Sales:      "AL",
		Percentage: "AM",
		Conversion: "AN",
		Separator:  "AO",
	}
	reportSkeleton["august"] = monthColumns{
		Listings:   "AP",
		Sales:      "AQ",
		Percentage: "AR",
		Conversion: "AS",
		Separator:  "AT",
	}
	reportSkeleton["september"] = monthColumns{
		Listings:   "AU",
		Sales:      "AV",
		Percentage: "AW",
		Conversion: "AX",
		Separator:  "AY",
	}
	reportSkeleton["october"] = monthColumns{
		Listings:   "AZ",
		Sales:      "BA",
		Percentage: "BB",
		Conversion: "BC",
		Separator:  "BD",
	}
	reportSkeleton["november"] = monthColumns{
		Listings:   "BE",
		Sales:      "BF",
		Percentage: "BG",
		Conversion: "BH",
		Separator:  "BI",
	}
	reportSkeleton["december"] = monthColumns{
		Listings:   "BJ",
		Sales:      "BK",
		Percentage: "BL",
		Conversion: "BM",
		Separator:  "BN",
	}
}
