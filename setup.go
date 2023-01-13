package main

import (
	"flag"
	"strconv"
	"strings"
	"time"
)

var option string
var store string
var month string
var year string

func setup() {
	loadFlags()
	processDates()
}

func loadFlags() {
	flag.StringVar(&option, "o", "none", "options: all/one")
	flag.StringVar(&year, "y", "none", "The Specific year to process")
	flag.StringVar(&month, "m", "none", "The Specific month to process")
	flag.StringVar(&store, "s", "none", "a specific store")
	flag.Parse()
}

func processDates() {
	curMonth := time.Now().Month() + 1
	if month == "none" {
		month = strings.ToLower(curMonth.String())
	}
	if year == "none" {
		year = strconv.Itoa(time.Now().Year())
	}
}
