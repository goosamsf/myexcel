package main

import (
	"fmt"
	"os"
	"time"
	_ "reflect"
	"strconv"
	"bufio"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/term"
	"golang.org/x/crypto/bcrypt"
)

/* Constants */
const GENERATE = 2
const RESET = 1
const FILESTATUS = 0

const FILEMAC = "/PTW_templates.xlsx"
const FILEWIN = "\\PTW_templates.xlsx"
const PASSWD = "$2a$10$SUDQmJzF0CZKxH8YDfomg.5BwVa2yYy7jYv6qMm84Ntgc1Ynya4bO"
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
	if initial_message(PASSWD) == 1 {
		press_enter_exit()
	}
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
	sheetList := f.GetSheetList()

	for {
		opt := chooseOption()
		switch opt {
		case FILESTATUS:
			checkFileStatus(f)
		case RESET:
			resetTemplate(f,sheetList)
		case GENERATE:
			fmt.Println("Process Begin...")
		default:
			press_enter_exit()
		}

		if opt == 2 {
			break
		}

	}	

	/* Get Requester Name */
	requesterNameHandler(f)
	
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

	fmt.Println("ALL WORK DONE. Program Gracefully TERMINATE")
	fmt.Println("Thank YOU")
	press_enter_exit()
}

func ceilingHandler(f *excelize.File, i int , row []string) {
	/* This function handles Ceiling Permit
		 First Generates the new sheet by copying template file 
		 and fill out necessary part with right value
	*/
	
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
	/* SAVE */
	f.Save()
}

func requesterNameHandler(f *excelize.File) {
	/*   This function handles to fill out the requester name in right
	  	 location 
	*/
	reader := bufio.NewReader(os.Stdin)
	fmt.Println("Type Requester Name: ")
	requesterName, _ := reader.ReadString('\n')
	requesterName = strings.TrimSpace(requesterName)

	f.SetCellValue(coldWork , "J43", requesterName)
	f.SetCellValue(coldWork , "CA65", requesterName)
	f.Save()
	fmt.Println("Press Enter to Proceed. ")
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

	// what is the name of the sheet it should be?
	sheetName := row[workLocation]
	sheetName = sheetName[:4] + strconv.Itoa(i) 

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

func prog_terminator() {
	fmt.Println("Terminating Program...")
	os.Exit(0)
}


func resetTemplate(f *excelize.File, sheetList []string) {
	/* Reseting Template... */
	if len(sheetList) <= 5 {
		fmt.Println("Template is already ready to process..  ")
		fmt.Println("Restart the program and Go ahead with Option 2...")
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
	fmt.Println("Restart the program and Go ahead with Option 2...")
	fmt.Println("시트 초기화 처리 완료.")
	fmt.Println("2번으로 생성 진행요망.")

	fmt.Println("Press Enter to Proceed...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
	f.Save()

}

func press_enter_exit() {
	/* In Window Environment, this function is necessary 
		otherwise user can't see any output and cmd windows is closed 
	*/
	fmt.Println("Press Enter to Exit...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
	os.Exit(0)
}

func initial_message(pwcheck string) int {

	fmt.Println("======================================================")
	fmt.Println("This Property belongs to SK C&C USA Infra Department")
	fmt.Println("Unauthorized use, copying, or distribution is strictly")
	fmt.Println("prohibited")
	fmt.Println("SK C&C USA Infra Department. All rights reserved.")
	fmt.Println("======================================================")
	fmt.Println("")

	fmt.Print("Password: ")
	bytePassword, err := term.ReadPassword(int(os.Stdin.Fd()))
	if err != nil {
		fmt.Println("\nError reading password:")
		press_enter_exit()
	}
	fmt.Println("Password Received, Validating User... ")
	userInput := string(bytePassword)	

	err = bcrypt.CompareHashAndPassword([]byte(pwcheck), []byte(userInput))		
	if err != nil {
		/* WRONG PASSWORD */ 
		fmt.Println("Failed to validate the PASSWORD( hint : m5 )")
		return 1	
	}else {
		fmt.Println("Successfully Validated")	
	}

	fmt.Println("Welcome to PTW Generator Program.")
	fmt.Println("This program generates next week's PTW based on \"Work Location\" sheet in \"PTW_tempalte.xlsx\" ")
	return 0
}

func chooseOption() int {
	reader := bufio.NewReader(os.Stdin) 
		 fmt.Print(`
===============================================
Choose an Option!
----------------------------------------------
0. File Status (생성가능여부 확인)
1. Reset Template (엑셀 템플릿 리셋)
2. Start Generating PTW (작업허가서 생성시작)
3. Exit (종료)
===============================================
`)
	 input , _ := reader.ReadString('\n')
	 input = strings.TrimSpace(input)
	 inputNum , err := strconv.Atoi(input)
	 if err != nil {
		 fmt.Println("Invalid number input")
		 press_enter_exit()
	 }
	 return inputNum
}

func checkFileStatus(f *excelize.File) {
	sheetList := f.GetSheetList()
	if len(sheetList) == 5 {
		fmt.Println(" SHHET 초기화 상태 확인")
		fmt.Println(" 2번으로 작업허가서 생성 시작요망")
	}else {
		fmt.Println(" SHEET 초기화 필요 ")
		fmt.Println(" 1번으로 SHEET 리셋 요망")
	}
	fmt.Println("Press Enter to Proceed...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
}
