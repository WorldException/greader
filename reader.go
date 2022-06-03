package main

import (
	"flag"
	"fmt"
	"greader/writers"
	"log"
	"os"
	"path"
	"strings"

	"github.com/extrame/xls"
	"github.com/tealeg/xlsx"
)

func read_xls(filename string, filepath string, fileresult string) error {
	if wb, err := xls.Open(filename, "utf-8"); err == nil {
		log.Printf("Author:%s", wb.Author)

		var fsheets *os.File
		fsheets, err = os.Create(fmt.Sprintf("%s.v8", fileresult))
		if err != nil {
			return err
		}
		defer fsheets.Close()

		for snum := 0; snum < wb.NumSheets(); snum++ {
			sname := wb.GetSheet(snum)
			filesheet := fmt.Sprintf("%s.v8_%d", fileresult, snum)
			log.Printf("Sheet: Num:%d; Rows:%d; Name:%s", snum, sname.MaxRow, sname.Name)
			// записать имя таблицы
			fsheets.WriteString(sname.Name + "\n")

			var v8 *writers.V8Table
			if v8, err = writers.NewV8Table(filesheet); err == nil {
				var cols []string = nil
				var rows []string = nil
				var skippedRows = 0
				for rnum := 0; rnum <= int(sname.MaxRow); rnum++ {
					row := sname.Row(rnum)
					FirstCol := row.FirstCol()
					LastCol := row.LastCol()
					if rnum < 10 {
						log.Printf("Row:%d, FirstCol:%d, LastCol:%d", rnum, FirstCol, LastCol)
					}
					if skippedRows > 10 {
						break
					}
					if FirstCol == LastCol {
						skippedRows++
						continue
					}
					if cols == nil {
						cols = []string{}
						for i := 0; i < LastCol; i++ {
							cols = append(cols, fmt.Sprintf("k_%d", i))
						}
						v8.AddHeader(cols)
					}
					if cols != nil {
						rows = []string{}
						for i := FirstCol; i < LastCol; i++ {
							rows = append(rows, strings.TrimSpace(row.Col(i)))
						}
						v8.AddRow(rows)
					}
				}
				v8.Close()
			} else {
				log.Printf("Error V8Table New:%s", err.Error())
			}

		}
	} else {
		return err
	}
	return nil
}

func read_xlsx(filename string, filepath string, fileresult string) error {
	if wb, err := xlsx.OpenFile(filename); err == nil {
		for snum := 0; snum < len(wb.Sheets); snum++ {
			sname := wb.Sheets[snum]
			log.Printf("Sheet: N:%d; R:%d; C:%d; Name:%s", snum, sname.MaxRow, sname.MaxCol, sname.Name)

			filesheet := fmt.Sprintf("%s.v8_%d", fileresult, snum)
			log.Printf("Sheet: Num:%d; Rows:%d; Name:%s", snum, sname.MaxRow, sname.Name)

			var v8 *writers.V8Table
			if v8, err = writers.NewV8Table(filesheet); err == nil {
				var cols []string = nil
				var rows []string = nil
				var skippedRows = 0
				for rnum := 0; rnum <= int(sname.MaxRow); rnum++ {
					row := sname.Row(rnum)
					FirstCol := 0
					LastCol := len(row.Cells)
					if rnum < 10 {
						log.Printf("Row:%d, FirstCol:%d, LastCol:%d", rnum, FirstCol, LastCol)
					}
					if skippedRows > 10 {
						break
					}
					if FirstCol == LastCol {
						skippedRows++
						continue
					}
					if cols == nil {
						cols = []string{}
						for i := 0; i < LastCol; i++ {
							cols = append(cols, fmt.Sprintf("k_%d", i))
						}
						v8.AddHeader(cols)
					}
					if cols != nil {
						rows = []string{}
						for i := FirstCol; i < LastCol; i++ {
							cell := row.Cells[i]
							switch cell.Type() {
							case xlsx.CellTypeDate:
								rows = append(rows, cell.String())
							default:
								rows = append(rows, strings.TrimSpace(cell.String()))
							}
						}
						v8.AddRow(rows)
					}
				}
				v8.Close()
			} else {
				log.Printf("Error V8Table New:%s", err.Error())
			}
		}
	} else {
		return err
	}

	return nil
}

func main() {
	var err error

	filename := flag.String("i", "", "Эксель файл")
	fileresult := flag.String("o", "", "Путь к файлу результата")
	flag.Parse()
	if len(*filename) == 0 {
		flag.Usage()
		log.Fatalln("Не указано имя файла -i filename")
	}
	var filepath string
	if len(*fileresult) == 0 {
		filepath = path.Dir(*filename)
		*fileresult = path.Join(filepath, "gresult")
	} else {
		filepath = path.Dir(*fileresult)
	}

	log.Printf("File: %s", *filename)
	log.Printf("Filepath: %s", filepath)
	log.Printf("result: %s", *fileresult)

	if strings.HasSuffix(strings.ToLower(*filename), ".xlsx") {
		err = read_xlsx(*filename, filepath, *fileresult)
	}
	if strings.HasSuffix(strings.ToLower(*filename), ".xls") {
		err = read_xls(*filename, filepath, *fileresult)
	}
	if err != nil {
		log.Printf("Error:%s", err.Error())
		os.Exit(1)
	}
	log.Print("end")
}
