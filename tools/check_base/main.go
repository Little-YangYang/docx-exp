package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	f, err := os.Open("examples/html_injection/template.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	fi, _ := f.Stat()

	// Check zip content
	zr, err := zip.NewReader(f, fi.Size())
	if err != nil {
		panic(err)
	}
	for _, zf := range zr.File {
		fmt.Println(zf.Name)
		if zf.Name == "word/styles.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			s := string(data)
			fmt.Println("Styles found:")
			parts := strings.Split(s, "<w:style ")
			for i, p := range parts {
				if i == 0 {
					continue
				}
				// Extract styleId
				idStart := strings.Index(p, "w:styleId=\"")
				if idStart == -1 {
					continue
				}
				idStart += 11
				idEnd := strings.Index(p[idStart:], "\"")
				if idEnd == -1 {
					continue
				}
				id := p[idStart : idStart+idEnd]

				// Extract name
				name := ""
				nameStart := strings.Index(p, "<w:name w:val=\"")
				if nameStart != -1 {
					nameStart += 15
					nameEnd := strings.Index(p[nameStart:], "\"")
					if nameEnd != -1 {
						name = p[nameStart : nameStart+nameEnd]
						fmt.Printf("ID: %s, Name: %s\n", id, name)
					}
				}
			}
		}
	}
}
