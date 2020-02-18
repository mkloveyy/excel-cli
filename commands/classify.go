package commands

import (
	"errors"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func Classify(filePath string, inputFileName string, sheetName string, outputFilePath string, column string) (err error) {
	if len(column) > 1 {
		return errors.New("wrong column value")
	}

	var f *excelize.File

	if f, err = excelize.OpenFile(filePath + inputFileName); err != nil {
		return
	}

	// get sheet
	sRows := f.GetRows(sheetName)
	// get sheet title
	sTitle := sRows[0]
	// get sheet data list
	sRows = sRows[1:]

	//total := len(sRows)

	// get the specific column and remove duplicate values
	var columnList []string

	for i := 2; i < len(sRows); i++ {
		isDuplicate := false
		value := f.GetCellValue(sheetName, column+strconv.Itoa(i))

		for _, j := range columnList {
			if j == value {
				isDuplicate = true
				break
			}
		}

		if !isDuplicate {
			columnList = append(columnList, value)
		}
	}

	fmt.Println(len(columnList))

	// make sure output file path exists
	if err = os.MkdirAll(filePath+outputFilePath, os.ModePerm); err != nil {
		return
	}

	// classify by diff column, file is named by column value
	for _, columnValue := range columnList {
		newFile := excelize.NewFile()
		newSheet := newFile.NewSheet(sheetName)
		// set title
		newFile.SetSheetRow(sheetName, "A1", &sTitle)

		k := 2

		for j := 2; j < len(sRows); j++ {
			if f.GetCellValue(sheetName, column+strconv.Itoa(j)) == columnValue {
				newFile.SetSheetRow(sheetName, "A"+strconv.Itoa(k), &sRows[j])
				k++
			}
		}

		newFile.SetActiveSheet(newSheet)

		if err = newFile.SaveAs(filePath + outputFilePath + columnValue + ".xlsx"); err != nil {
			return
		}
	}

	return
}
