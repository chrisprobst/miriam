package main

import (
	"log"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	path := "/Users/chrisprobst/Desktop/Fz.xlsx"
	file, err := xlsx.OpenFile(path)
	if err != nil {
		panic(err)
	}

	for _, sheet := range file.Sheets {
		log.Print(sheet.Name)

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

			tableNameString, err := tableName.String()
			if err != nil {
				panic(err)
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

				// Do something with these data...
				log.Printf("%s -> %s -> %s", tableNameString, tableBeta1MeanString, tableBeta3MeanString)
			}

			rowOffset += rowStep
		}

		// for _, row := range sheet.Rows {
		// 	if len(row.Cells) == 0 {
		// 		continue
		// 	}
		//
		// 	if name, err := row.Cells[0].String(); err == nil && strings.Contains(name, "Baseline") {
		// 		log.Print(name)
		// 	}
		//
		// 	// for _, cell := range row.Cells {
		// 	//
		// 	// }
		// }
	}

	file.Save("/Users/chrisprobst/Desktop/F4_copy.xlsx")
}
