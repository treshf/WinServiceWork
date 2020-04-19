package main

import (
	"fmt"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (
	sheet   = "Sheet1"
	frmDate = "02.01.2006"
	frmTime = "15:4:5"
)

//StyleCentre Получение стиля по центру я чейки
func StyleCentre(xlsx *excelize.File) (int, error) {
	return xlsx.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical:   "center",
			Horizontal: "center",
		},
	})
}

//SetText Ширина столбца по тексту ячейки
func SetText(cell string, text string, xlsx *excelize.File) error {
	err := xlsx.SetColWidth(sheet,
		fmt.Sprintf("%c", cell[0]), fmt.Sprintf("%c", cell[0]), float64(len(text)))
	if err != nil {
		return err
	}
	return xlsx.SetCellStr(sheet, cell, text)
}

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

	style, err := StyleCentre(xlsx)
	if err != nil {
		return fmt.Errorf("Ошибка стиля:%v", err)
	}

	for i, d := range top {
		cell, err := excelize.CoordinatesToCellName(i+1, 1)
		if err != nil {
			return fmt.Errorf("Ошибка преобразования координаты ячейки:%v", err)
		}

		if err = SetText(cell, d, xlsx); err != nil {
			return fmt.Errorf("немогу записать в ячейку:%v", err)
		}
		if err = xlsx.SetCellStyle(sheet, cell, cell, style); err != nil {
			return fmt.Errorf("Ошибка задания стиля:%v", err)
		}
	}

	cell, _ := excelize.CoordinatesToCellName(len(top), 2)

	if err = xlsx.SetCellFormula(sheet, cell, "=G2-F2"); err != nil {
		return fmt.Errorf("Ошибка записи в ячейку:%s %v", "R2C8", err)
	}

	if err = xlsx.SaveAs(name); err != nil {
		return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
	}

	return nil
}

//FileName генерирует имя файла по текущему месяцу и году
func FileName() (string, error) {
	_, pathExe, err := exePath()
	if err != nil {
		return "", fmt.Errorf("Не могу получить путь к exe. Error: %v", err)
	}
	name := fmt.Sprintf(pathExe+"\\%d.%d.xlsx", time.Now().Month(), time.Now().Year())
	return name, nil
}

//AddDataStart добавление информации о новом дне
func AddDataStart(row int, xlsx *excelize.File) error {
	var err error
	err = SetText("A"+fmt.Sprintf("%d", row), time.Now().Format(frmDate), xlsx)
	if err != nil {
		return fmt.Errorf("Ошибка записи в ячейку xlsx:%v", err)
	}
	err = xlsx.SetCellValue(sheet, "B"+fmt.Sprintf("%d", row), time.Now().UTC())
	if err != nil {
		return fmt.Errorf("Ошибка записи в ячейку xlsx:%v", err)
	}
	val, err := xlsx.GetCellValue(sheet, "B"+fmt.Sprintf("%d", row))
	if err != nil {
		return fmt.Errorf("Ошибка чтения из ячейки xlsx:%v", err)
	}
	xlsx.SetColWidth(sheet, "B", "B", float64(len(val))*1.2)
	return nil
}

//Start добавляет информацию о начале рабочего дня
func Start() error {
	name, err := FileName()
	if err != nil {
		return fmt.Errorf("Ошибка создания имени файла: %v", err)
	}

	if _, err := os.Stat(name); err != nil {
		if err := InitFIle(name); err != nil {
			return fmt.Errorf("Невозможно инициализировать файл:%s %v", name, err)
		}
	}

	xlsx, err := excelize.OpenFile(name)
	if err != nil {
		return fmt.Errorf("Ошибка открытия файла:%s %v", name, err)
	}
	rows, err := xlsx.GetRows("Sheet1")
	if err != nil {
		return fmt.Errorf("не могу прочитать строки:%s %v", name, err)
	}

	if len(rows[len(rows)-1][0]) > 0 {
		dateXLSX, err := time.Parse(frmDate, rows[len(rows)-1][0])
		if err != nil {
			return fmt.Errorf("Ошибка чтения даты:%s %v", rows[len(rows)-1][0], err)
		}
		if dateXLSX.Day() < time.Now().Day() {
			err = AddDataStart(len(rows)+1, xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}

			if err = xlsx.SaveAs(name); err != nil {
				return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
			}
		}
	} else {
		err = AddDataStart(len(rows), xlsx)
		if err != nil {
			return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
		}
		if err = xlsx.SaveAs(name); err != nil {
			return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
		}
	}
	return nil
}

//End добавляет информацию о конце рабочего дня
func End() error {

	return nil
}
