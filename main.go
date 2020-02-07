package main

import (
	"log"
	"os"
	"time"

	"github.com/urfave/cli/v2"

	"excel/commands"
)

func main() {
	var path, filePath, fileName, sheetName string

	var length int

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
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "path",
				Aliases:     []string{"p"},
				Value:       "/Users/mklyy/Desktop/",
				Usage:       "root path of pending files",
				Destination: &path,
			},
		},
		Commands: []*cli.Command{
			{
				Name:    "split",
				Aliases: []string{"sp"},
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
					return commands.Split(path, fileName, sheetName, filePath, length)
				},
			},
			{
				Name:    "unify",
				Aliases: []string{"un"},
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
		},
	}

	if err := excelCli.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}
