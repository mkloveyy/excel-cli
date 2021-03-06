package main

import (
	"log"
	"os"
	"time"

	"github.com/urfave/cli/v2"

	"excel/commands"
)

func main() {
	var rootPath string

	var filePath, fileName, sheetName string

	var column string

	var length int

	var ownerColumn, companyColumn string

	excelCli := &cli.App{
		Name:     "excel-cli",
		Version:  "v1.0.0",
		Compiled: time.Now(),
		Authors: []*cli.Author{
			{
				Name:  "Mark Ma",
				Email: "ma_k@ctrip.com",
			},
		},
		Copyright: "(c) Authored By Mark Ma",
		Usage:     "some useful actions for excel files",
		// todo - why not work on windows?
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "root-path",
				Aliases:     []string{"rp"},
				Value:       "",
				Usage:       "root path of pending files",
				Destination: &rootPath,
			},
		},
		Commands: []*cli.Command{
			{
				Name:    "split",
				Aliases: []string{"s"},
				Usage:   "split sheet into multiple files",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "file-name",
						Aliases:     []string{"fn"},
						Value:       "test.xlsx",
						Usage:       "name of pending files to be split",
						Destination: &fileName,
					},
					&cli.StringFlag{
						Name:        "sheet-name",
						Aliases:     []string{"sn"},
						Value:       "数据源",
						Usage:       "name of file sheet to be split",
						Destination: &sheetName,
					},
					&cli.StringFlag{
						Name:        "sub-path",
						Aliases:     []string{"sp"},
						Value:       "/",
						Usage:       "sub path of file to output",
						Destination: &filePath,
					},
					&cli.IntFlag{
						Name:        "length",
						Aliases:     []string{"l"},
						Value:       100,
						Usage:       "length of each sub file sheet",
						Destination: &length,
					},
				},
				Action: func(c *cli.Context) error {
					return commands.Split(rootPath, fileName, sheetName, filePath, length)
				},
			},
			{
				// todo - unify mutiple files into one sheet
				Name:    "unify",
				Aliases: []string{"u"},
				Usage:   "unify sheets in multiple files into one sheet",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "sub-path",
						Aliases:     []string{"sp"},
						Value:       "/",
						Usage:       "sub path of files to be unified",
						Destination: &filePath,
					},
					&cli.StringFlag{
						Name:        "file-name",
						Aliases:     []string{"fn"},
						Value:       "test.xlsx",
						Usage:       "name of file to output",
						Destination: &fileName,
					},
				},
				Action: func(c *cli.Context) error {
					return commands.Unify(filePath, fileName)
				},
			},
			{
				Name:    "classify",
				Aliases: []string{"c"},
				Usage:   "classify data of sheet into multiple files with different value in a column",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "file-name",
						Aliases:     []string{"fn"},
						Value:       "test.xlsx",
						Usage:       "name of pending files to be split",
						Destination: &fileName,
					},
					&cli.StringFlag{
						Name:        "sheet-name",
						Aliases:     []string{"sn"},
						Value:       "数据源",
						Usage:       "name of file sheet to be split",
						Destination: &sheetName,
					},
					&cli.StringFlag{
						Name:        "sub-path",
						Aliases:     []string{"sp"},
						Value:       "/",
						Usage:       "sub path of file to output",
						Destination: &filePath,
					},
					&cli.StringFlag{
						Name:        "column",
						Aliases:     []string{"c"},
						Value:       "A",
						Usage:       "name of the column",
						Destination: &column,
					},
				},
				Action: func(c *cli.Context) error {
					return commands.Classify(rootPath, fileName, sheetName, filePath, column)
				},
			},
			{
				Name:    "correct",
				Aliases: []string{"ct"},
				Usage:   "correct owner of company",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "file-name",
						Aliases:     []string{"fn"},
						Value:       "test.xlsx",
						Usage:       "name of pending files to be split",
						Destination: &fileName,
					},
					&cli.StringFlag{
						Name:        "sheet-name",
						Aliases:     []string{"sn"},
						Value:       "数据源",
						Usage:       "name of file sheet to be split",
						Destination: &sheetName,
					},
					&cli.StringFlag{
						Name:        "owner-column",
						Aliases:     []string{"oc"},
						Value:       "",
						Usage:       "column of owner",
						Destination: &ownerColumn,
					},
					&cli.StringFlag{
						Name:        "company-column",
						Aliases:     []string{"cc"},
						Value:       "",
						Usage:       "column of company",
						Destination: &companyColumn,
					},
				},
				Action: func(c *cli.Context) error {
					return commands.Correct(rootPath, fileName, sheetName, ownerColumn, companyColumn)
				},
			},
		},
	}

	if err := excelCli.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}
