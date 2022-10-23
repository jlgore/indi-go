package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"

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
		for _, colCell := range row {
			data = append(data, []string{string(colCell[B]), string(colCell[C]), string(colCell[D]), string(colCell[F])})
		}
	}

	fmt.Println(data)

}
