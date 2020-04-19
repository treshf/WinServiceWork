package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
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

	if err = xlsx.SetCellFormula(sheet, "H2", "=G2-F2"); err != nil {
		return fmt.Errorf("Ошибка записи в ячейку:%s %v", "H2", err)
	}

	if err = xlsx.SetCellFormula(sheet, "F2", "=(СУММ(D2:D50) * 24) + E2*8"); err != nil {
		return fmt.Errorf("Ошибка записи в ячейку:%s %v", "F2", err)
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
func AddDataStart(row string, timeIn time.Time, xlsx *excelize.File) error {
	var err error
	err = SetText("A"+row, time.Now().Format(frmDate), xlsx)
	if err != nil {
		return fmt.Errorf("Ошибка записи в ячейку xlsx:%v", err)
	}
	err = SetCellTime("B"+row, timeIn, xlsx)
	if err != nil {
		return err
	}
	return nil
}

//AddDataEnd добавление информации об окончании рабочего дня
func AddDataEnd(row string, timeIn time.Time, xlsx *excelize.File) error {
	var err error
	err = SetCellTime("C"+row, timeIn, xlsx)
	if err != nil {
		return err
	}
	err = xlsx.SetCellFormula(sheet, "D"+row, "=B"+row+"-C"+row+"-\"1:00\"")
	if err != nil {
		return err
	}
	return nil
}

//SetCellTime записывает в ячейку время и выравнивает ширину столбца по содержимому
func SetCellTime(cell string, timeIn time.Time, xlsx *excelize.File) error {
	var err error
	err = xlsx.SetCellValue(sheet, cell, timeIn.UTC())
	if err != nil {
		return fmt.Errorf("Ошибка записи в ячейку xlsx:%v", err)
	}
	val, err := xlsx.GetCellValue(sheet, cell)
	if err != nil {
		return fmt.Errorf("Ошибка чтения из ячейки xlsx:%v", err)
	}
	col := fmt.Sprintf("%c", cell[0])
	columnW, _ := xlsx.GetColWidth(sheet, col)
	if columnW < float64(len(val))*1.2 {
		xlsx.SetColWidth(sheet, col, col, float64(len(val))*1.2)
	}
	return nil
}

//OpenXLSX открывает файл и возвращает строки и объект для записи
func OpenXLSX() ([][]string, *excelize.File, string, error) {
	name, err := FileName()
	if err != nil {
		return nil, nil, "", fmt.Errorf("Ошибка создания имени файла: %v", err)
	}

	if _, err := os.Stat(name); err != nil {
		if err := InitFIle(name); err != nil {
			return nil, nil, "", fmt.Errorf("Невозможно инициализировать файл:%s %v", name, err)
		}
	}

	xlsx, err := excelize.OpenFile(name)
	if err != nil {
		return nil, nil, "", fmt.Errorf("Ошибка открытия файла:%s %v", name, err)
	}
	rows, err := xlsx.GetRows("Sheet1")
	if err != nil {
		return nil, nil, "", fmt.Errorf("не могу прочитать строки:%s %v", name, err)
	}
	return rows, xlsx, name, nil
}

//GetTimeStandart вычисляет время по производственному календарю
func GetTimeStandart() (int, error) {
	_, pathExe, err := exePath()
	if err != nil {
		return 0, fmt.Errorf("Не могу получить путь к exe. Error: %v", err)
	}
	xlsx, err := excelize.OpenFile(pathExe + "\\Пр. календарь " + fmt.Sprint(time.Now().Year()) + ".xlsx")

	starty := 5
	startx := 3
	nowMouth := int(time.Now().Month())
	for {
		if nowMouth > 3 {
			nowMouth -= 3
			starty += 8
		} else {
			startx += (nowMouth - 1) * 6
			break
		}
	}

	yBuf := starty
	xBuf := startx
	syleMap := make(map[int]int)
	var styles []int
	for {
		cell, _ := excelize.CoordinatesToCellName(xBuf, yBuf)
		str, err := xlsx.GetCellValue("Календарь", cell)
		if err != nil {
			return 0, fmt.Errorf("Ошибка чтения ячейки %s: %v", cell, err)
		}
		if len(str) > 0 {
			if i, _ := strconv.Atoi(str); i == time.Now().Day() {
				break
			}
			style, _ := xlsx.GetCellStyle("Календарь", cell)
			_, ok := syleMap[style]
			if ok {
				syleMap[style]++
			} else {
				syleMap[style] = 1
				styles = append(styles, style)
			}
		}
		yBuf++

		if yBuf == starty+7 {
			yBuf = starty
			xBuf++
			if xBuf == startx+6 {
				break
			}
		}
	}

	if len(styles) > 3 {
		return 0, fmt.Errorf("Стилей больше 3: %s", fmt.Sprint(styles))
	}

	sort.Ints(styles)

	hour := syleMap[styles[len(styles)-1]] * 8
	if len(styles) > 2 {
		hour += syleMap[styles[0]] * 7
	}

	return hour, nil
}

//StartWork добавляет информацию о начале рабочего дня
func StartWork() error {
	rows, xlsx, name, err := OpenXLSX()
	if err != nil {
		return fmt.Errorf("Start: %v", err)
	}

	if len(rows[len(rows)-1][0]) > 0 {
		dateXLSX, err := time.Parse(frmDate, rows[len(rows)-1][0])
		if err != nil {
			return fmt.Errorf("Ошибка чтения даты:%s %v", rows[len(rows)-1][0], err)
		}
		if dateXLSX.Day() != time.Now().Day() {
			err = AddDataStart(fmt.Sprint(len(rows)+1), time.Now(), xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}

			if err = xlsx.SaveAs(name); err != nil {
				return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
			}
		}
	} else {
		err = AddDataStart(fmt.Sprint(len(rows)), time.Now(), xlsx)
		if err != nil {
			return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
		}
		if err = xlsx.SaveAs(name); err != nil {
			return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
		}
	}

	return nil
}

//EndWork добавляет информацию о конце рабочего дня
func EndWork() error {
	rows, xlsx, name, err := OpenXLSX()
	if err != nil {
		return fmt.Errorf("Start: %v", err)
	}

	rowStr := fmt.Sprint(len(rows))
	if len(rows[len(rows)-1][0]) > 0 {
		dateXLSX, err := time.Parse(frmDate, rows[len(rows)-1][0])
		if err != nil {
			return fmt.Errorf("Ошибка чтения даты:%s %v", rows[len(rows)-1][0], err)
		}
		if dateXLSX.Day() == time.Now().Day() {
			err = AddDataEnd(rowStr, time.Now(), xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}
		} else {
			log.Output(1, "Не правильное время окончания. Возможно PC не был выключен")
			bufTime, _ := time.Parse(frmTime, "22:00:0")
			err = AddDataEnd(rowStr, bufTime, xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}

			rowStr = fmt.Sprint(len(rows) + 1)
			bufTime, _ = time.Parse(frmTime, "10:00:0")
			err = AddDataStart(rowStr, bufTime, xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}

			err = AddDataEnd(rowStr, time.Now(), xlsx)
			if err != nil {
				return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
			}
		}
	} else {
		log.Output(1, "Запись о завершении без начала")
		err = AddDataEnd(rowStr, time.Now(), xlsx)
		if err != nil {
			return fmt.Errorf("Ошибка записи данных:%s %v", name, err)
		}
	}

	timeStandart, err := GetTimeStandart()
	if err != nil {
		return fmt.Errorf("Ошибка загрузки стандартов: %v", err)
	}

	if err = xlsx.SetCellInt(sheet, "G2", timeStandart); err != nil {
		return fmt.Errorf("Ошибка записи int %d %v", timeStandart, err)
	}

	if err = xlsx.SaveAs(name); err != nil {
		return fmt.Errorf("Ошибка сохранения xlsx:%s %v", name, err)
	}

	return nil
}
