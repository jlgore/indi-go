package main

import (
	"fmt"
	"os"

	"github.com/olekukonko/tablewriter"
	"github.com/xuri/excelize/v2"
	//"bufio"
)

func main() {
	f, err := excelize.OpenFile("SEC510 Index.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	for index, name := range f.GetSheetMap() {
		data := [][]string{}
		fmt.Println(index)
		rows, err := f.GetRows(name)
		if err != nil {
			fmt.Println(err)
			return
		}

		for i, row := range rows {
			if i == 0 {
				continue
			}
			data = append(data, []string{row[1], row[2], row[3], row[4]})
		}

		fileName := fmt.Sprintf("SEC510 Index %v.md", name)

		mkdoc, err := os.OpenFile(fileName, os.O_WRONLY|os.O_CREATE|os.O_APPEND, 0644)
		if err != nil {
			fmt.Println(err)
			return
		}
		defer mkdoc.Close()

		header := fmt.Sprintf("# %v\n", name)

		pageHeader, err := mkdoc.WriteString(header)
		if err != nil {
			panic(err)
			fmt.Println(pageHeader)
		}

		table := tablewriter.NewWriter(mkdoc)
		table.SetAutoWrapText(false)
		table.SetHeader([]string{"Page", "Slide Title", "Keyword 1", "Keyword 2"})
		table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
		table.SetCenterSeparator("|")
		table.AppendBulk(data)
		table.Render()

	}

}
