package writers

import (
	"bufio"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
)

const (
	cHeadPattern    = "{\"#\",acf6192e-81ca-46ef-93a6-5a6968b78663,\n{8,\n{%d,\n{-2,\"НомерСтроки\",\n{\"Pattern\",\n{\"N\"}\n},\"N\",10}"
	cHeadColPattern = ",\n{%d,\"%s\",\n{\"Pattern\",\n{\"S\"}\n},\"%s\",10}"
)

type V8Table struct {
	filename string
	columns  []string
	w        *bufio.Writer
	f        *os.File
	flines   *os.File
	lineNo   int
}

func NewV8Table(filename string) (*V8Table, error) {
	v := V8Table{
		filename: filename,
		lineNo:   0,
	}
	log.Printf("V8 file: %s", v.filename)
	var err = *new(error)
	v.f, err = os.Create(filename)
	v.w = bufio.NewWriter(v.f)
	v.flines, err = ioutil.TempFile(os.TempDir(), "lines")
	if err != nil {
		return &v, err
	}
	return &v, err
}

func (v *V8Table) AddHeader(cols []string) error {
	var err error = nil
	fmt.Fprintf(v.w, cHeadPattern, len(cols))

	var headTail = "0,-2"
	for i, _ := range cols {
		headTail += fmt.Sprintf("%d,%d,", i+1, i)
	}

	for i, name := range cols {
		fmt.Fprintf(v.w, cHeadColPattern, i, name, name)
	}
	fmt.Fprint(v.w, "},\n")
	fmt.Fprintf(v.w, "{2,%d,\n%s", len(cols), headTail)

	return err
}

func (v *V8Table) AddRow(cols []string) error {
	// line header
	if v.lineNo > 0 {
		fmt.Fprint(v.flines, "0},\n")
	}
	fmt.Fprintf(v.flines, "{2,%d,%d,\n", v.lineNo, len(cols))
	fmt.Fprintf(v.flines, "{\"N\",%d},", v.lineNo+1)
	for _, value := range cols {
		fmt.Fprintf(v.flines, "\n{\"S\",\"%s\"},", value)
	}
	v.lineNo += 1
	return nil
}

func (v *V8Table) tail() {
	fmt.Fprintf(v.w, "\n{1,%d,\n", v.lineNo)
	v.flines.Seek(0, 0)
	io.Copy(v.w, v.flines)
	fmt.Fprintf(v.w, "\n},-1,%d}\n}\n}", v.lineNo-1)
}

func (v *V8Table) Close() {
	if v.w != nil {
		v.tail()
		v.w.Flush()
	}
	if v.f != nil {
		v.f.Close()
	}
	if v.flines != nil {
		v.flines.Close()
		log.Printf("Remove temp:%s", v.flines.Name())
		os.Remove(v.flines.Name())
	}
}
