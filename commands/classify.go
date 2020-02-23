package commands

import (
	"errors"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type File struct {
	file  *excelize.File
	count int
}

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

	// Output some message
	fmt.Println(fmt.Sprintf("file count: %d", len(columnList)))
	fmt.Println("column list: ", columnList)

	// make sure output file path exists
	if err = os.MkdirAll(filePath+outputFilePath, os.ModePerm); err != nil {
		return
	}

	// cache all null files, file is named by column value
	outFiles := make(map[string]File)

	for _, columnValue := range columnList {
		newFile := excelize.NewFile()
		newSheet := newFile.NewSheet(sheetName)
		// set title
		newFile.SetSheetRow(sheetName, "A1", &sTitle)
		newFile.SetActiveSheet(newSheet)

		outFiles[columnValue] = File{file: newFile, count: 2}
	}

	// classify by diff column
	for j := 2; j < len(sRows); j++ {
		// init used data
		columnValue := f.GetCellValue(sheetName, column+strconv.Itoa(j))
		file := outFiles[columnValue].file
		count := outFiles[columnValue].count

		// set row to specific file
		outFiles[columnValue].file.SetSheetRow(sheetName, "A"+strconv.Itoa(count), &sRows[j-1])

		// update file obj
		count++
		outFiles[columnValue] = File{file: file, count: count}
	}

	// save all sub files
	for c, f := range outFiles {
		if err = f.file.SaveAs(filePath + outputFilePath + c + ".xlsx"); err != nil {
			return
		}
	}

	return
}
