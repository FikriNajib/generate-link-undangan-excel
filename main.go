package main

import (
	"fmt"
	"log"
	"strings"
	"net/url"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func generateURL(name string) string {
	escapedName := url.QueryEscape(name)
	urlTemplate := "https://invitadosproject.com/fikri-fatimah/?to=%s"
	return fmt.Sprintf(urlTemplate, strings.ReplaceAll(escapedName, "+", "&"))
}

func main() {
	// Input Excel file path
	inputFilePath := "fikri.xlsx"

	// Output Excel file path
	outputFilePath := "output.xlsx"

	// Open the input Excel file
	xlsx, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		log.Fatal(err)
	}

	// Get all rows from the "Sheet1" sheet
	rows := xlsx.GetRows("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// Create a new Excel file for output
	outputXLSX := excelize.NewFile()

	// Create a new sheet in the output file
	// outputSheet := outputXLSX.NewSheet("Sheet1")

	// Iterate over the rows and generate URLs
	for rowIndex, row := range rows {
		// Skip the header row (assuming the first row is the header)
		if rowIndex == 0 {
			continue
		}

		// Extract the name from the current row
		if len(row) > 0 {
			name := row[0]

			// Generate the URL for the current name
			url := generateURL(name)

			// Write the name and URL to the output sheet
			outputXLSX.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowIndex), name)
			outputXLSX.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowIndex), url)
		}
	}

	// Save the output Excel file
	if err := outputXLSX.SaveAs(outputFilePath); err != nil {
		log.Fatal(err)
	}
}
