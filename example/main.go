package main

import (
	"github.com/fumiama/go-docx"
	"github.com/vc2402/templdocx"
	"os"
	"text/template"
)

var values = map[string]interface{}{
	"title":      "Using templdocx library",
	"userName":   "John Doe",
	"createDate": "2025-05-01",
	"events": []any{
		map[string]any{
			"time":        "2025-06-02 12:23",
			"description": "John made some action",
		},
		map[string]any{
			"time":        "2025-06-02 12:25",
			"description": "Mr Doe made another action",
		},
	},
}

var funcMap = template.FuncMap{
	"description": func() string {
		return `We can build result values in functions
or use existing. It depends on our needs. For sure we can
pass params to functions`
	},
	"images": func() []any {
		return []any{
			map[string]any{"name": "first", "imageName": "a.jpg"},
			map[string]any{"name": "second", "imageName": "b.jpg"},
		}
	},
}

func main() {

	var err error

	readFile, err := os.Open("Template.docx")
	if err != nil {
		panic(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)
	if err != nil {
		panic(err)
	}
	dp := templdocx.NewDocParser(
		doc,
		values,
		funcMap,
		templdocx.CustomControlProvider(
			func(name string) templdocx.ControlObject {
				if name == "image" {
					return &templdocx.ControlTypeDrawing{}
				}
				return nil
			},
		),
	)

	err = dp.Process()
	if err != nil {
		panic(err)
	}
	f, err := os.Create("generated.docx")
	// save to file
	if err != nil {
		panic(err)
	}
	_, err = doc.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
}
