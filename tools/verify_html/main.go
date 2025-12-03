package main

import (
	"fmt"
	"os"

	"github.com/fumiama/go-docx"
)

func main() {
	f, err := os.Open("examples/row_if/result_true.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	doc, err := docx.Parse(f, fi.Size())
	if err != nil {
		panic(err)
	}

	fmt.Println("--- Items ---")
	for _, it := range doc.Document.Body.Items {
		if p, ok := it.(*docx.Paragraph); ok {
			style := ""
			if p.Properties != nil && p.Properties.Style != nil {
				style = p.Properties.Style.Val
			}
			fmt.Printf("P [%s]: ", style)
			printP(p)
			fmt.Println()
		} else if t, ok := it.(*docx.Table); ok {
			fmt.Println("Table:")
			for _, row := range t.TableRows {
				for _, cell := range row.TableCells {
					fmt.Print("  Cell: ")
					for _, p := range cell.Paragraphs {
						style := ""
						if p.Properties != nil && p.Properties.Style != nil {
							style = p.Properties.Style.Val
						}
						fmt.Printf("P [%s]: ", style)
						printP(p)
						fmt.Println()
					}
					fmt.Println()
				}
			}
		}
	}
}

func printP(p *docx.Paragraph) {
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			for _, runChild := range run.Children {
				if text, ok := runChild.(*docx.Text); ok {
					fmt.Print(text.Text)
				} else if _, ok := runChild.(*docx.Drawing); ok {
					fmt.Print("[IMAGE]")
				}
			}
		}
	}
}
