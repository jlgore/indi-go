package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"github.com/olekukonko/tablewriter"
	"os"
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

	rows, err := f.GetRows("Book 1")
	if err != nil {
		fmt.Println(err)
		return
	}

	data := [][]string{}

	for _, row := range rows {
		//fmt.Println(row[3])	
		data = append(data, []string{row[1], fmt.Sprintf("%s\n", row[2]), row[3], row[4]})
		//fmt.Printf("Page: %s, Slide Title: %s, Keyword 1: %s\n", row[1], row[2], row[3])
	}	

	mkdoc, err := os.OpenFile("SEC510-Index.md", os.O_WRONLY|os.O_CREATE|os.O_APPEND, 0644)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer mkdoc.Close()
	
	table := tablewriter.NewWriter(mkdoc)
	table.SetAutoWrapText(false)
	table.SetHeader([]string{"Page", "Slide Title", "Keyword 1", "Keyword 2"})
	table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
	table.SetCenterSeparator("|")
	table.AppendBulk(data)
	table.Render()

}
