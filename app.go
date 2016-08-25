package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////// Files //////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

const (
	fzPath     = "/Users/chrisprobst/Desktop/Fz.xlsx"
	f4Path     = "/Users/chrisprobst/Desktop/F4.xlsx"
	inputPath  = "/Users/chrisprobst/Desktop/EEG_Auswertung.xlsx"
	outputPath = "/Users/chrisprobst/Desktop/EEG_Auswertung_generated.xlsx"
)

var (
	fzFile    *xlsx.File
	f4File    *xlsx.File
	inputFile *xlsx.File
)

func openFiles() {
	var err error

	if fzFile, err = xlsx.OpenFile(fzPath); err != nil {
		panic(err)
	}

	if f4File, err = xlsx.OpenFile(f4Path); err != nil {
		panic(err)
	}

	if inputFile, err = xlsx.OpenFile(inputPath); err != nil {
		panic(err)
	}
}

func saveOutputFile() {
	if err := inputFile.Save(outputPath); err != nil {
		panic(err)
	}
}

////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////// Extract ////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

func copyFromFToOutput(category string, f *xlsx.File) {
	for _, sheet := range f.Sheets {

		// Parse person id
		personID, err := strconv.ParseInt(sheet.Name[2:4], 10, 32)
		if err != nil {
			panic(err)
		}

		//////////////////////////////////////////////////////
		///////////////// Extract F values ///////////////////
		//////////////////////////////////////////////////////
		rowOffset := 5
		rowBeta1Offset := 7
		rowBeta3Offset := 8
		rowStep := 12
		for {
			// Skip incomplete rows
			if len(sheet.Rows) < rowOffset {
				break
			}

			// Extract Beta1 and Beta3
			tableName := sheet.Rows[rowOffset].Cells[0]
			tableBeta1Mean := sheet.Rows[rowOffset+rowBeta1Offset].Cells[2]
			tableBeta3Mean := sheet.Rows[rowOffset+rowBeta3Offset].Cells[2]

			// Check and map table name
			tableNameString, err := tableName.String()
			if err != nil {
				panic(err)
			}
			tableNameString = strings.Split(strings.Split(tableNameString, ":")[1], "(")[0]
			tableNameString = tableNameString[:len(tableNameString)-1]
			tableNameString = strings.Replace(tableNameString, ".", "", 1)
			if strings.Contains(tableNameString, "Baseline") {
				tableNameString = strings.Replace(tableNameString, "Baseline", "B", 1)
			}

			// We do not need any Cue values...
			if !strings.Contains(tableNameString, "Cue") {

				tableBeta1MeanString, err := tableBeta1Mean.String()
				if err != nil {
					panic(err)
				}

				tableBeta3MeanString, err := tableBeta3Mean.String()
				if err != nil {
					panic(err)
				}

				// Define LB column and cell value
				lbColName := fmt.Sprintf("%sLB_%s", category, tableNameString)
				lbCellValue, err := strconv.ParseFloat(strings.TrimSpace(tableBeta1MeanString), 64)
				if err != nil {
					panic(err)
				}

				// Define HB column and cell value
				hbColName := fmt.Sprintf("%sHB_%s", category, tableNameString)
				hbCellValue, err := strconv.ParseFloat(strings.TrimSpace(tableBeta3MeanString), 64)
				if err != nil {
					panic(err)
				}

				// Do something with these data...
				log.Printf("[VP%d] %s: %f", personID, lbColName, lbCellValue)
				log.Printf("[VP%d] %s: %f", personID, hbColName, hbCellValue)

				// Insert into sheet
				insertCellIntoOutput(personID, lbColName, lbCellValue)
				insertCellIntoOutput(personID, hbColName, hbCellValue)
			}

			rowOffset += rowStep
		}
	}
}

////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////// Insert /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

func insertCellIntoOutput(personID int64, col string, cell float64) {
	mainSheet := inputFile.Sheets[0]
	columnNames := mainSheet.Rows[0].Cells

	// First find row index
	for i, row := range mainSheet.Rows[1:] {
		if len(row.Cells) == 0 {
			continue
		}

		if parsedPersonID, err := strconv.ParseInt(row.Cells[0].Value, 10, 32); err == nil && parsedPersonID == personID {
			for j, colName := range columnNames {
				if colName.Value != col {
					continue
				}

				i++

				mainSheet.Cell(i, j).SetFloat(cell)
				return
			}

			log.Printf("[VP%d] Could not find column: %s", personID, col)
			return
		}
	}

	log.Printf("[VP%d] Could not find person id: %d", personID, personID)
}

func main() {
	openFiles()
	copyFromFToOutput("F4", f4File)
	copyFromFToOutput("FZ", fzFile)
	saveOutputFile()
}
