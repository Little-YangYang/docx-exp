package main

import (
	"fmt"
	"os"

	"github.com/fumiama/go-docx"
	docxexp "github.com/little-yangyang/docx-exp"
)

func main() {
	f, err := os.Open("examples/complex_report/template.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	tpl, err := docxexp.New(f, fi.Size())
	if err != nil {
		panic(err)
	}

	data := map[string]interface{}{
		"ProjectName": "Test Project",
		"Vulns": []map[string]interface{}{
			{"Name": "Vuln 1", "Severity": "High"},
			{"Name": "Vuln 2", "Severity": "Medium"},
		},
		"Image": docxexp.ImageInjector{
			Path:   "testdata/test_image.png",
			Width:  100,
			Height: 100,
		},
	}

	if err := tpl.Render(data); err != nil {
		panic(err)
	}

	out, err := os.Create("examples/complex_report/result.docx")
	if err != nil {
		panic(err)
	}
	defer out.Close()

	if err := tpl.Save(out); err != nil {
		panic(err)
	}
	fmt.Println("Rendered examples/complex_report/result.docx")
}

func createComplexTemplate() {
	f, err := os.Open("testdata/base.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	doc, err := docx.Parse(f, fi.Size())
	if err != nil {
		panic(err)
	}

	doc.Document.Body.Items = []interface{}{}

	doc.AddParagraph().AddText("Project: {{.ProjectName}}")

	// Table loop
	table := doc.AddTable(2, 2, 9000, nil)
	table.TableRows[0].TableCells[0].AddParagraph().AddText("Name")
	table.TableRows[0].TableCells[1].AddParagraph().AddText("Severity")

	// Row loop
	row := table.TableRows[1]
	row.TableCells[0].AddParagraph().AddText("{{ range .Vulns }}{{ .Name }}")
	row.TableCells[1].AddParagraph().AddText("{{ .Severity }}")

	doc.AddParagraph().AddText("Image:")
	doc.AddParagraph().AddText("{{ inject .Image }}")

	out, err := os.Create("examples/complex_report/template.docx")
	if err != nil {
		panic(err)
	}
	defer out.Close()
	doc.WriteTo(out)
}
