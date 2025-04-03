package main


import (
    "fmt"
		"os"
		"time"
    "github.com/xuri/excelize/v2"
)


func main() {
	/* Open Excel File */
	f, err := excelize.OpenFile("PTW_templates.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	date_handler(f)
	
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sh_list := f.GetSheetList()
	fmt.Println(sh_list)
	for i:= 0; i < len(sh_list); i++ {
		if i > 5 {
			f.DeleteSheet(sh_list[i])
		}
	}
	index,_ := f.NewSheet("NewSheet2")
	err = f.CopySheet(2, index)
	if err != nil {
			fmt.Println("CopySheet error:", err)
	}
	f.Save()
	os.Exit(1)

	rows, err := f.GetRows("Work Location")
	rows = rows[1:]
	if err != nil {
			fmt.Println(err)
			return
	}
	/* Iterate over entire row */
	/* Main Loop */
	for i, row := range rows{
		// i is index  : int
		// row is current row : []string
		//fmt.Printf("Type of _row_ : %T\n",row)
		//fmt.Printf("Laying out %dth row\n",i)
		
		// put together name 
		cold_handler(f,i, row)
		// Handle ptw_cold
		// Handle Risk Assessment
		// Handle CEILING 
		// Handle 

			
		if i == 5 {
			break
		}	
	}

	


}

func date_handler(f *excelize.File) {
	/* Get current day(mon thru sun) , get current date  subtract that value
	* from 7 */
	fmt.Println(time.Now())	
	os.Exit(0)


}

func cold_handler(f *excelize.File ,i int , row []string) {
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

