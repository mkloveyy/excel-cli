package commands

import (
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func Split(filePath string, inputFileName string, sheetName string, outputFilePath string, length int) (err error) {
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

	total := len(sRows)

	if err = os.MkdirAll(filePath+outputFilePath, os.ModePerm); err != nil {
		return
	}

	for i := 0; i <= total/length; i++ {
		newFile := excelize.NewFile()
		newSheet := newFile.NewSheet(sheetName)

		// set title
		newFile.SetSheetRow(sheetName, "A1", &sTitle)

		k := 2

	generateOneFile:
		for j := i * length; j < (i+1)*length; j++ {
			// to the end
			if j > total-1 {
				break generateOneFile
			}

			// set data
			newFile.SetSheetRow(sheetName, "A"+strconv.Itoa(k), &sRows[j])
			k++
		}

		newFile.SetActiveSheet(newSheet)

		if err = newFile.SaveAs(filePath + outputFilePath + strconv.Itoa(i) + ".xlsx"); err != nil {
			return
		}
	}

	return
}
