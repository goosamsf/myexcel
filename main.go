package main

import (
	"fmt"
	"os"
	"time"
	_ "reflect"
	"strconv"
	"bufio"
	"runtime"

	"github.com/xuri/excelize/v2"
)

/* Constants */
const FILEMAC = "/PTW_templates.xlsx"
const FILEWIN = "\\PTW_templates.xlsx"
const coldWorkIndex = 1
const riskAssIndex = 3
const ceilingIndex = 4
const cntToSat = 5
const cntToSun = 6
const cntToWeekAfter = 7

const coldWork = "PTW_COLD"
const ceilingWork = "Above Ceiling"
const riskAssessment = "Risk Assessment"

const validDateBegin = "S34"
const validDateEnd = "AJ35"
const closingDateCell = "AZ132"
const (
	// This name is used to access each row's column
	workLocation int = iota
	needBit
	ceilingBit
	elec_certBit
	workDesc_1
	workDesc_2
	toolUsed
	ra_workDesc
	haz_1
	mit_1
	haz_2
	mit_2
	haz_3
	mit_3
)

var glob_cw int

/* -------------------------------------- */
/* --					M 	A 	I		N						--- */
/* -------------------------------------- */
func main() {
	// VAR
	initial_message()
	press_enter_exit()

	var template_path string
	cwd , err := os.Getwd()
	if err != nil {
		fmt.Println("Err : ", err)
		return
	}
	running := runtime.GOOS

	/* 	-- OS CHECK -- */
	switch running {
	case "darwin":
		// MAC OS
		cwd = cwd + FILEMAC
		template_path = cwd
	case "windows":
		// WINDOWS
		cwd = cwd + FILEWIN
		template_path = cwd
	}

	i := 1
	ceil_cnt := 0

	/* Open Excel File */
	f, err := excelize.OpenFile(template_path)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Reseting Template 
  sheetList := f.GetSheetList()
	if len(sheetList) > 5 {
		resetTemplate(f ,sheetList)
	}


	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	/*  Handles Date in Coldwork / Ceiling / Risk Assessment */
	dateHandler(f)

	/* Save What has been Done */
	f.Save()

	/* Get Rows */
	rows, err := f.GetRows("Work Location")
	rows = rows[1:]
	
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(" ---- PROGRAM BEGIN ----")
	/* Iterate over entire row */
				/* Main Loop */
	for _ , row := range rows {
		
		if row[needBit] == "0" {
			continue
		}
		fmt.Printf("PTW : %s..\n", row[workLocation])

		// COLD WORK HANDLER ( MOST OF THE WORK DONE HERE ) 
		cold_handler(f, i, row)
		// ---------------------------------------------- //

		if row[ceilingBit] == "1" {
			// If ceiling is needed it handles ceiling 
			ceilingHandler(f,i, row)
			ceil_cnt += 1
		}

		i += 1

	}
	fmt.Println(" ---- SUMMARY ---- ")
	fmt.Printf("%d PTW(s): Done.\n", i)
	fmt.Printf("%d RISK ASSESSMENT(s): Done.\n", i)
	fmt.Printf("%d Ceiling PERMIT: Done.\n", ceil_cnt)

	press_enter_exit()

}

func ceilingHandler(f *excelize.File, i int , row []string) {
	// Ceiling Handler 
	sheetName := row[workLocation]

	workLoc := sheetName
	sheetName = sheetName[:4] + strconv.Itoa(i) 
	sheetName = "CEIL_" + sheetName
	index, _ := f.NewSheet(sheetName)
	err := f.CopySheet(ceilingIndex, index) 
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}
	f.SetCellValue(sheetName, "D13",workLoc )
	f.Save()
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

	/* --- PTW COLD --- */ 
	f.SetCellValue(coldWork, validDateBegin, validFrom)
	f.SetCellValue(coldWork, validDateEnd, validUntil)
	f.SetCellValue(coldWork, closingDateCell, closingDate)

	/* --- Ceiling Work --- */
	f.SetCellValue(ceilingWork, "D10" ,dateForSign)    	// Application Date 
	f.SetCellValue(ceilingWork, "L10" ,validFrom) 		 	// Valid Date begin
	f.SetCellValue(ceilingWork, "U10" ,validUntil) 		 	// Valid Date end
	f.SetCellValue(ceilingWork, "W46" ,dateForSign )   	// HE Signature
	f.SetCellValue(ceilingWork, "W56" ,dateForSign )   	// SHE Signature
	f.SetCellValue(ceilingWork, "W63" ,validFrom) 			// Performing authority

	/* --- Risk Assessment --- */
	f.SetCellValue(riskAssessment, "D8" , validFrom)    // Performing authority


	f.Save()
	bufio.NewReader(os.Stdin).ReadBytes('\n')

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
	sheetName := row[workLocation]
	sheetName = sheetName[:4] + strconv.Itoa(i) 

	// what is the content that should be filled in permit registry number

	index, _ := f.NewSheet(sheetName)
	err := f.CopySheet(1, index) // 
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}

	// The cell number for permit registry is AK9
	permitRegistry := "SKCCUS_" + sheetName + "_CW" + strconv.Itoa(glob_cw)
	permitRegCell := "AK9"

	f.SetCellValue(sheetName, permitRegCell ,permitRegistry) // Permit Registry Number
	f.SetCellValue(sheetName, "J22" , row[workLocation]) // Work Location 
	f.SetCellValue(sheetName, "K24" , row[workDesc_1])   // Work Description 1
	f.SetCellValue(sheetName, "K28" , row[workDesc_2]) 	 // Work Description 2
	f.SetCellValue(sheetName, "Q31" , row[toolUsed]) 	   // Tool Used
	f.Save()

	ra_handler(f, permitRegistry, row, i)

}


func ra_handler(f *excelize.File , permitRegistry string , row []string, i int){
	/* 	-- Function Description -- 
		 This function handles risk assessmen sheet.
		 First it create a sheet by copying risk assessment sheet. 
		 Then, it fills out necessary cell such as permitRegistry no , work Area
	*/
	sheetName := row[workLocation]
	sheetName = sheetName[:4] + strconv.Itoa(i) 
	sheetName = "RA_" + sheetName
	index, _ := f.NewSheet(sheetName)
	err := f.CopySheet(riskAssIndex, index) 
	if err != nil {
		fmt.Println("CopySheet error:", err)
	}

	f.SetCellValue(sheetName , "E3", permitRegistry )
	f.SetCellValue(sheetName , "D7", row[workLocation])
	
	ra_workDescCell := 21

	for ind := 0; ind < 3; ind++ {
		wd := "B"  //Cell Column for Work Description
		wd = wd + strconv.Itoa(ra_workDescCell)

		f.SetCellValue(sheetName, wd, row[ra_workDesc])
		ra_workDescCell += 2
	}

	/* 		--- Risk Assessment Table Fill out --- */
	f.SetCellValue(sheetName, "D21", row[haz_1])
	f.SetCellValue(sheetName, "G21", row[mit_1])
	f.SetCellValue(sheetName, "D23", row[haz_2])
	f.SetCellValue(sheetName, "G23", row[mit_2])
	f.SetCellValue(sheetName, "D25", row[haz_3])
	f.SetCellValue(sheetName, "G25", row[mit_3])

	/* If You need to write something more in riskassesment sheet 
		 write it here */

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


func resetTemplate(f *excelize.File, sheetList []string) {
	/* Reseting Template... */
	if len(sheetList) <= 5 {
		fmt.Println("Template is ready, no need to go further processing.. ")
		prog_terminator()
	}

	fmt.Println("Reseting Template ...")

	for i , item := range sheetList{
		if i < 5 {
			// Index UPTO 4 is used as template. 
			continue
		}
		f.DeleteSheet(item)
	}
	fmt.Println("DONE.")
	fmt.Println("You can now Generate next week's PTW / Good luck. ")
	f.Save()

	prog_terminator()
}

func press_enter_exit() {
	/* In Window Environment, this function is necessary 
		otherwise user can't see any output and cmd windows is closed 
	*/
	fmt.Println("Press Enter to Exit...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
	os.Exit(0)
}

func initial_message() {
// ==============================================
// This program belongs to ABC Corporation.
// Unauthorized use, copying, or distribution
// is strictly prohibited.
// © 2025 ABC Corporation. All rights reserved.
// ==============================================
	fmt.Println("======================================================")
	fmt.Println("This Property belongs to SK C&C USA Infra Department")
	fmt.Println("Unauthorized use, copying, or distribution is strictly")
	fmt.Println("prohibited")
	fmt.Println("SK C&C USA Infra Department. All rights reserved.")
	fmt.Println("======================================================")
	fmt.Println("")
	fmt.Println("Welcome to PTW Generator Program.")
	fmt.Println("This program generates next week's PTW based on \"Work Location\" sheet in \"PTW_tempalte.xlsx\" ")
	

}

