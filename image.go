package docxexp

import (
	"github.com/fumiama/go-docx"
)

// ImageInjector injects an image into the document
type ImageInjector struct {
	Path   string
	Width  int64
	Height int64
}

// Inject implements the Injector interface
func (i ImageInjector) Inject(doc *docx.Docx, p *docx.Paragraph) ([]interface{}, error) {
	// Add drawing
	// p.AddInlineDrawingFrom takes a file path
	_, err := p.AddInlineDrawingFrom(i.Path)
	if err != nil {
		return nil, err
	}

	// Append run to paragraph children
	// p.Children = append(p.Children, run) // Redundant
	return nil, nil
}
