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
const cntToSat = 5
const cntToSun = 6
const cntToWeekAfter = 7

const coldWork = "PTW_COLD"
const validDateBegin = "S34"
const validDateEnd = "AJ35"
const closingDateCell = "AZ132"

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
	//nextMonday := getComingMonday(f)
	nextMonday := dateHandler(f)
	fmt.Println(nextMonday)

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
	

	index, _ := f.NewSheet("NewSheet2")
	err = f.CopySheet(2, index)
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}

	//prog_terminator()

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

func dateHandler(f *excelize.File) int {
	/* This functon handles all the dates in the excel file
		 date should be written in following cells:
		 - Valid Date Begin & End
		 - Date for Approval & Issue
		 - Date for Daily Inspection 
	 */
	var dateArray [6]string
	
	// Cell Location for Approval & Issue
	var appAndIssue = [4]string{"BG65", "BG70", "BG75", "BG80"}

	// Cell Location for Daily Inspections
	var dailyInspection = [6]string{"BA103", "BA105","BA107", "BX103","BX105", "BX107"}

	// Get current time.
	localnow := time.Now()
	weekday := int(localnow.Weekday())

	// Get #Days to the coming Monday
	dateToMonday := 7 - weekday + 1

	validFrom := localnow.AddDate(0,0,dateToMonday).Format("01/02")
	validUntil := localnow.AddDate(0,0,dateToMonday + cntToSat).Format("01/02")
	closingDate := localnow.AddDate(0,0,dateToMonday + cntToWeekAfter).Format("01/02")
	dateForSign := localnow.AddDate(0,0,dateToMonday-4).Format("01/02")

	for i := 0; i<len(dailyInspection); i++ {
	// Date that should be written in Daily Inspection
		dateArray[i] = localnow.AddDate(0,0,dateToMonday+i).Format("01/02")
		f.SetCellValue(coldWork, dailyInspection[i], dateArray[i])
	}

	for i:= 0; i < len(appAndIssue); i++ {
	// Date that should be written in Approval & Issue
		f.SetCellValue(coldWork, appAndIssue[i], dateForSign)
	}

	// Write 
	f.SetCellValue(coldWork, validDateBegin, validFrom)
	f.SetCellValue(coldWork, validDateEnd, validUntil)
	f.SetCellValue(coldWork, closingDateCell, closingDate)
	f.Save()

	return 0
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

func prog_terminator() {
	fmt.Println("Terminating Program...")
	os.Exit(1)
}
