package main

import (
	"fmt"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//InitFIle создаёт и наполняет новый файл
func InitFIle(name string) error {
	var err error
	xlsx := excelize.NewFile()

	top := []string{
		"Дата",
		"Начало (ч)",
		"Окончание (ч)",
		"За день (ч)",
		"Отпуск (д)",
		"За месяц (ч)",
		"Должно быть (ч)",
		"Доработать (ч)",
	}

	for i, d := range top {
		err = xlsx.SetCellStr("Sheet1", "R1C"+fmt.Sprintf("%d", i+1), d)
		if err != nil {
			return fmt.Errorf("немогу записать в ячейку:%s %v", "R1C"+fmt.Sprintf("%d", i+1), err)
		}
	}
	err = xlsx.SetCellStr("Sheet1", "R2C8", "=G2-F2")
	if err != nil {
		return fmt.Errorf("немогу записать в ячейку:%s %v", "R2C8", err)
	}
	err = xlsx.SaveAs(name)

	if err != nil {
		return fmt.Errorf("невозможно сохранить файл:%s %v", name, err)
	}
	return nil
}

//FileName генерирует имя файла по текущему месяцу и году
func FileName() string {
	name := fmt.Sprintf("%d.%d.xslx", time.Now().Month(), time.Now().Year())
	return name
}

//Start добавляет информацию о начале рабочего дня
func Start() error {
	name := FileName()
	//if _, err := os.Stat(name); err != nil {
	err := InitFIle(name)
	if err != nil {
		return fmt.Errorf("невозможно инициализировать файл:%s %v", name, err)
	}
	//}
	/*
		xlsx, err := excelize.OpenFile(name)
		if err != nil {
			return fmt.Errorf("не открыть файл:%s %v", name, err)
		}
		rows, err := xlsx.GetRows("Sheet1")
		if err != nil {
			return fmt.Errorf("не могу прочитать строки:%s %v", name, err)
		}

		for _, row := range rows {

		}
	*/
	return nil
}
