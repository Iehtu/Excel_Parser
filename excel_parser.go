package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("Реестр_требований-ФР-Доработок_БК_v1(Приложение 1)_010724.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	comments, err := f.GetComments("Лист1")
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, value := range comments {

		fmt.Println(value)
	}

}
