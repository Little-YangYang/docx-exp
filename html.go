package docxexp

import (
	"encoding/base64"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"

	"github.com/fumiama/go-docx"
	"golang.org/x/net/html"
)

// HTMLInjector injects HTML content into the document
type HTMLInjector struct {
	Content string
}

// Inject implements the Injector interface
func (h HTMLInjector) Inject(doc *docx.Docx, p *docx.Paragraph) ([]interface{}, error) {
	// Parse HTML
	node, err := html.Parse(strings.NewReader(h.Content))
	if err != nil {
		return nil, err
	}

	var items []interface{}

	// Traverse HTML nodes
	var f func(*html.Node)
	f = func(n *html.Node) {
		if n.Type == html.DocumentNode {
			for c := n.FirstChild; c != nil; c = c.NextSibling {
				f(c)
			}
		} else if n.Type == html.ElementNode {
			switch n.Data {
			case "h1", "h2", "h3":
				// Create new paragraph
				newP := createParagraph(doc)
				newP.XMLName = p.XMLName
				newP.Style(headingStyle(n.Data))

				// Extract text
				text := extractText(n)
				run := &docx.Run{
					Children: []interface{}{&docx.Text{Text: text}},
				}
				newP.Children = append(newP.Children, run)
				items = append(items, newP)
			case "p":
				newP := createParagraph(doc)
				newP.XMLName = p.XMLName
				// Handle children (text, img, etc)
				processChildren(doc, n, newP)
				items = append(items, newP)
			case "img":
				newP := createParagraph(doc)
				newP.XMLName = p.XMLName
				if err := addImage(doc, newP, n); err == nil {
					items = append(items, newP)
				}
			case "body", "html":
				for c := n.FirstChild; c != nil; c = c.NextSibling {
					f(c)
				}
			default:
				// Ignore other tags or treat as text?
				// If it's div or something, recurse.
				for c := n.FirstChild; c != nil; c = c.NextSibling {
					f(c)
				}
			}
		} else if n.Type == html.TextNode {
			// Top level text? Create paragraph.
			text := strings.TrimSpace(n.Data)
			if text != "" {
				newP := createParagraph(doc)
				newP.XMLName = p.XMLName
				run := &docx.Run{
					Children: []interface{}{&docx.Text{Text: text}},
				}
				newP.Children = append(newP.Children, run)
				items = append(items, newP)
			}
		}
	}
	f(node)
	return items, nil
}

func createParagraph(doc *docx.Docx) *docx.Paragraph {
	p := doc.AddParagraph()
	// Remove from doc items to avoid duplication, as we will add it to the list manually
	if len(doc.Document.Body.Items) > 0 {
		doc.Document.Body.Items = doc.Document.Body.Items[:len(doc.Document.Body.Items)-1]
	}
	return p
}

func headingStyle(tag string) string {
	switch tag {
	case "h1":
		return "1"
	case "h2":
		return "2"
	case "h3":
		return "3"
	default:
		return "Normal"
	}
}

func extractText(n *html.Node) string {
	var text string
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.TextNode {
			text += c.Data
		} else {
			text += extractText(c)
		}
	}
	return text
}

func processChildren(doc *docx.Docx, n *html.Node, p *docx.Paragraph) {
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.TextNode {
			run := &docx.Run{
				Children: []interface{}{&docx.Text{Text: c.Data}},
			}
			p.Children = append(p.Children, run)
		} else if c.Type == html.ElementNode && c.Data == "img" {
			addImage(doc, p, c)
		} else {
			processChildren(doc, c, p)
		}
	}
}

func addImage(doc *docx.Docx, p *docx.Paragraph, n *html.Node) error {
	var src string
	for _, attr := range n.Attr {
		if attr.Key == "src" {
			src = attr.Val
			break
		}
	}
	if src == "" {
		return fmt.Errorf("no src")
	}

	if strings.HasPrefix(src, "data:image/") {
		// Base64
		parts := strings.Split(src, ",")
		if len(parts) != 2 {
			return fmt.Errorf("invalid base64 image")
		}
		data, err := base64.StdEncoding.DecodeString(parts[1])
		if err != nil {
			return err
		}
		_, err = p.AddInlineDrawing(data)
		if err != nil {
			return err
		}
	} else if strings.HasPrefix(src, "http") {
		resp, err := http.Get(src)
		if err != nil {
			return err
		}
		defer resp.Body.Close()

		data, err := io.ReadAll(resp.Body)
		if err != nil {
			return err
		}

		_, err = p.AddInlineDrawing(data)
		if err != nil {
			return err
		}
	} else {
		// Local file
		// Verify file exists
		if _, err := os.Stat(src); err != nil {
			return err
		}

		_, err := p.AddInlineDrawingFrom(src)
		if err != nil {
			return err
		}
	}
	return nil
}
