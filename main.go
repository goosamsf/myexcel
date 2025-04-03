package main

import (
	"fmt"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

/* Constants */
const coldWorkIndex = 1
const riskAssIndex = 3
const ceilingIndex = 4

const coldWork = "PTW_COLD"
const validDateBegin = "S34"

/*
  f.GetSheetList() --> []string
	f.DeleteSheet(string) --> void





*/

func main() {
	/* Open Excel File */
	f, err := excelize.OpenFile("PTW_templates.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Handle Date in Coldwork
	nextMonday := getComingMonday(f)
	/* TODO
	getComingMonday function should handle more than what it does now.
	right now it just gets the date for incoming Monday but that's not enough.
	it should also know that following dates:
		- date for closing                 --> validDateBegin + WeekAfter
		- date for daily signed ( 6days)   --> validDateBegin + 1,2,3, ... 6
		- date for signed ( in step 1,2,3) --> validDateBegin - 4

	Then it should be manually embedded into the right cells so that I don't
	have to worry about cell'style and go with given style.

	*/

	f.Save()
	os.Exit(0)

	index, _ := f.NewSheet("NewSheet2")
	err = f.CopySheet(2, index)
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}
	os.Exit(1)

	rows, err := f.GetRows("Work Location")
	rows = rows[1:]
	if err != nil {
		fmt.Println(err)
		return
	}
	/* Iterate over entire row */
	/* Main Loop */
	for i, row := range rows {
		// i is index  : int
		// row is current row : []string
		//fmt.Printf("Type of _row_ : %T\n",row)
		//fmt.Printf("Laying out %dth row\n",i)

		// put together name
		cold_handler(f, i, row)
		// Handle ptw_cold
		// Handle Risk Assessment
		// Handle CEILING
		// Handle

		if i == 5 {
			break
		}
	}

}

func getComingMonday(f *excelize.File) string {
	/* This function returns the date of coming Monday as string in "mm/dd" form */
	localnow := time.Now()
	weekday := int(time.Now().Weekday())
	dateToMonday := 7 - weekday + 1

	// AddDate with dateToMondy in third argument should return coming Monday's Date
	// It will be returning as string with the date format : mm/dd
	// 01 represents month / 02 represents date
	return localnow.AddDate(0, 0, dateToMonday).Format("01/02")

}

func cold_handler(f *excelize.File, i int, row []string) {
	/* Generate sheet that contains cold ptw */
	/*
		1. Create empty sheet
		2. copy ptw work to the empty sheet
		3. Work Location
		4. Work Description
	*/
}

func read_field_names(f *excelize.File) {
	// create a map

	// find a way to iterate over first column until
	// value is null

	return

}
