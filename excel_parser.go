package main

import (
	"errors"
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

const outFileName string = "Book1.xlsx"
const inFileName string = "Реестр_требований-ФР-Доработок_БК_v1(Приложение 1)_010724.xlsx"

func main() {

	var f_out *excelize.File

	f_out = nil

	f, err := excelize.OpenFile(inFileName)
	if err != nil {
		fmt.Println(err)
		return
	}

	if _, err := os.Stat(outFileName); err == nil {

		f_out, err = excelize.OpenFile(outFileName)
		if err != nil {
			fmt.Println(err)
			return
		}

	} else if errors.Is(err, os.ErrNotExist) {

		f_out = excelize.NewFile()

	} else {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	defer func() {
		if f_out != nil {
			if err := f_out.Close(); err != nil {
				fmt.Println(err)
			}
		}

	}()
	comments, err := f.GetComments("Лист1")
	if err != nil {
		fmt.Println(err)
		return
	}

	counter := 2

	for _, value := range comments {

		f_out.SetCellValue("Лист1", "A1", "Номер ячейки")
		f_out.SetCellValue("Лист1", "B1", "Текст ячейки")
		f_out.SetCellValue("Лист1", "C1", "Комментарий")

		val, _ := f.CalcCellValue("Лист1", value.Cell)

		f_out.SetCellValue("Лист1", "A"+strconv.Itoa(counter), value.Cell)
		f_out.SetCellValue("Лист1", "B"+strconv.Itoa(counter), val)
		f_out.SetCellRichText("Лист1", "C"+strconv.Itoa(counter), value.Paragraph)
		counter++
	}
	f_out.SaveAs(outFileName)

}
