package main

import (
	"fmt"
	"os"
	"time"
	"reflect"

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

const (
	workLocation int = iota
	needbit
	ceilingbit
	elec_certbit
	workDesc_1
	workDesc_2
	tool_used
)

var glob_cw int

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
	nextMonday := dateHandler(f)
	fmt.Println(nextMonday)

	// Save What has been Done
	f.Save()


	rows, err := f.GetRows("Work Location")
	rows = rows[1:]
	fmt.Println("Type of rows:", reflect.TypeOf(rows))
	
	//prog_terminator()

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
		prog_terminator()
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

	// ----  Warning Global Variable is used
	_ , cw := localnow.ISOWeek()
	glob_cw = cw+1
	// Gloval Variable warning ----

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
	// what is the name of the sheet it should be?
	

	index, _ := f.NewSheet("NewSheet2")
	err := f.CopySheet(2, index)
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}
	f.Save()
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


func tester_GetRows(rows [][]string) {
	fmt.Println(rows)
	prog_terminator()
}

