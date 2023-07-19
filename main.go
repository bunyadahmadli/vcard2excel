package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strings"
)

// Define a structure to hold vCard data
type vCard struct {
	FirstName string
	Phone     string
}

func main() {
	// Replace "input.vcf" with the path to your vCard file
	vcardFilePath := "vcard.vcf"

	// Open the vCard file
	file, err := os.Open(vcardFilePath)
	if err != nil {
		fmt.Println("Error opening vCard file:", err)
		return
	}
	defer file.Close()

	// Create a new Excel file
	f := excelize.NewFile()
	index, err := f.NewSheet("Sheet1")
	f.SetActiveSheet(index)

	// Set up the Excel headers
	headers := []string{"First Name", "Phone"}
	for col, header := range headers {
		f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", 'A'+col, 1), header)
	}

	// Starting row number in Excel
	rowNum := 2

	// Parse the vCard data
	scanner := bufio.NewScanner(file)
	var vc vCard

	for scanner.Scan() {
		line := scanner.Text()

		if strings.HasPrefix(line, "FN:") {
			vc.FirstName = strings.TrimPrefix(line, "FN:")
		} else if strings.Contains(line, "TEL") {
			names := strings.Split(line, ":")
			if len(names) >= 2 {
				vc.Phone = names[1]
				continue
			}
			vc.Phone = strings.TrimPrefix(line, "TEL;")
		}

		// If we have collected all the necessary info for a contact
		if vc.FirstName != "" && vc.Phone != "" {
			// Write the data to Excel
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowNum), vc.FirstName)

			f.SetCellValue("Sheet1", fmt.Sprintf("D%d", rowNum), vc.Phone)

			rowNum++

			// Reset the vCard for the next contact
			vc = vCard{}
		}
	}

	// Save the Excel file
	excelFilePath := "output.xlsx" // Replace with the desired output path
	if err := f.SaveAs(excelFilePath); err != nil {
		fmt.Println("Error saving Excel file:", err)
		return
	}

	fmt.Println("vCard converted to Excel successfully.")
}
