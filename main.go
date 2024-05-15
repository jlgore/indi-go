package main

import (
	"bytes"
	"fmt"
	"os"

	wkhtmltopdf "github.com/SebastiaanKlippert/go-wkhtmltopdf"
	"github.com/olekukonko/tablewriter"
	"github.com/xuri/excelize/v2"
	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/renderer/html"
	//"bufio"
)

func main() {
	f, err := excelize.OpenFile("Index SEC541.xlsx")
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
			rowData := []string{}
			for j := 1; j <= 4; j++ {
				if j < len(row) {
					rowData = append(rowData, row[j])
				} else {
					rowData = append(rowData, "")
				}
			}
			data = append(data, rowData)
		}

		fileName := fmt.Sprintf("Index SEC541 %v.md", name)

		mkdoc, err := os.OpenFile(fileName, os.O_WRONLY|os.O_CREATE|os.O_APPEND, 0644)
		if err != nil {
			fmt.Println(err)
			return
		}
		defer mkdoc.Close()

		header := fmt.Sprintf("# %v\n", name)

		_, err = mkdoc.WriteString(header)
		if err != nil {
			fmt.Println(err)
			return
		}

		// Create a new table writer
		table := tablewriter.NewWriter(mkdoc)
		table.SetHeader([]string{"PAGE", "SLIDE TITLE", "KEYWORD 1", "KEYWORD 2"})
		table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
		table.SetCenterSeparator("|")
		table.SetAutoWrapText(false)
		table.SetAlignment(tablewriter.ALIGN_LEFT)

		// Append the data rows to the table
		for _, row := range data {
			table.Append(row)
		}

		// Render the table
		table.Render()

		// Read the Markdown file
		mdFile, err := os.ReadFile(fileName)
		if err != nil {
			fmt.Println(err)
			return
		}

		// Convert markdown to HTML
		var buf bytes.Buffer
		md := goldmark.New(
			goldmark.WithExtensions(extension.Table),
			goldmark.WithRendererOptions(
				html.WithUnsafe(),
			),
		)
		if err := md.Convert(mdFile, &buf); err != nil {
			fmt.Println(err)
			return
		}
		html := buf.String()

		// Create a new PDF generator
		pdfg, err := wkhtmltopdf.NewPDFGenerator()
		if err != nil {
			fmt.Println(err)
			return
		}
		// Set the HTML content with custom CSS styles
		htmlWithStyles := fmt.Sprintf(`
    <html>
    <head>
        <style>
            table {
                border-collapse: collapse;
                width: 100%%;
            }
            th, td {
                border: 1px solid black;
                padding: 8px;
                word-wrap: break-word;
                white-space: normal;
            }
        </style>
    </head>
    <body>
        %s
    </body>
    </html>
`, html)
		pdfg.AddPage(wkhtmltopdf.NewPageReader(bytes.NewReader([]byte(htmlWithStyles))))
		// Set the PDF file path
		pdfFileName := fmt.Sprintf("Index SEC541 %v.pdf", name)

		// Create a new file
		file, err := os.Create(pdfFileName)
		if err != nil {
			fmt.Println(err)
			return
		}
		defer file.Close()

		// Set the output to the file
		pdfg.SetOutput(file)

		// Generate the PDF file
		if err := pdfg.Create(); err != nil {
			fmt.Println(err)
			return
		}
	}
}
