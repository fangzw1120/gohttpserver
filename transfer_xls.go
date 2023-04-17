package main

import (
	"errors"
	"fmt"
	interLogger "git.woa.com/forisfang_tut/logger"
	"github.com/Lofanmi/chinese-calendar-golang/calendar"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
	"time"
)

// User 不考虑闰月问题
var period = 10
var users = make([]*User, 0)

type User struct {
	ID             string
	Name           string
	Year           int64
	Month          int64
	Day            int64
	Birthday       string
	DateType       string
	FutureBirthday map[int64]string
}

const (
	UserIDHeader       = "工号"
	UserNameHeader     = "姓名"
	UserBirthHeader    = "生日"
	UserDateTypeHeader = "日历类型"
)

func transferXLS(filePath, outputPath string) {
	//filePath := "/Users/forisfang/Desktop/Book1.xlsx"
	header, sheetName, err := readExcelFile(filePath)
	if err != nil {
		interLogger.Errorf("%+v", err)
	}
	users = GetAllUser()

	nowYear := time.Now().Year()
	totYears := period
	yearLT := make([]int64, totYears)
	yearStrLt := make([]string, totYears)
	for i := 0; i < totYears; i++ {
		val := nowYear + i
		yearLT[i] = int64(val)
		yearStrLt[i] = strconv.Itoa(val) + "年"
	}
	interLogger.Debugf("calculate period: %+v", yearLT)

	// for user
	cnt := 0
	for _, user := range users {
		cnt++
		interLogger.Infof("user: %+v", user)
		futureBirthday := make(map[int64]string, totYears)
		for _, year := range yearLT {
			if user.DateType == "阴历" {
				// 对应年份农历的时间对象
				valLunarItem := calendar.ByLunar(year, user.Month, user.Day, 0, 0, 0, false)
				// 对应年份农历
				//valLunar := valLunarItem.Lunar
				// 对应年份公历
				valSolar := valLunarItem.Solar
				futureBirthday[year] = fmt.Sprintf("%+v-%+v-%+v", valSolar.GetYear(), GetSolarMonthStr(valLunarItem), GetSolarDayStr(valLunarItem))
				//logger.Debugf("年份: %+v, 农历: %+v-%+v-%+v, Solar: %+v-%+v-%+v", year,
				//	valLunar.GetYear(), valLunar.GetMonthStr(), valLunar.GetDayStr(),
				//	valSolar.GetYear(), valSolar.GetMonthStr(), valSolar.GetDayStr())

			} else if user.DateType == "阳历" {
				month, day := "", ""
				if user.Month <= 9 {
					month = "0" + strconv.Itoa(int(user.Month))
				} else {
					month = strconv.Itoa(int(user.Month))
				}
				if user.Day <= 9 {
					day = "0" + strconv.Itoa(int(user.Day))
				} else {
					day = strconv.Itoa(int(user.Day))
				}
				futureBirthday[year] = fmt.Sprintf("%+v-%+v-%+v", year, month, day)
			}

		}
		user.FutureBirthday = futureBirthday
		interLogger.Debugf("user: %+v", user)
	}
	interLogger.Infof("tot user: %+v", cnt)

	newHeader := append(header, yearStrLt...)
	//writeToFile("/Users/forisfang/Desktop/Book1_result.xlsx", sheetName, newHeader, yearLT, users)
	writeToFile(outputPath, sheetName, newHeader, yearLT, users)

	//t := time.Now()
	//c := calendar.ByTimestamp(t.Unix())
	//lunarItem := c.Lunar
	//lunarStr := fmt.Sprintf("农历 %+v 年 %+v 月 %+v 日", lunarItem.GetYear(), lunarItem.GetMonth(), lunarItem.GetDay())
	//
	//bytes, err := c.ToJSON()
	//if err != nil {
	//	logger.Errorf("%+v", err)
	//}
	//logger.Debug(string(bytes))
	//logger.Debug(lunarStr)
}

func GetSolarMonthStr(l *calendar.Calendar) string {
	if l.Solar.GetMonth() <= 9 {
		return "0" + strconv.Itoa(int(l.Solar.GetMonth()))
	}
	return strconv.Itoa(int(l.Solar.GetMonth()))
}

func GetSolarDayStr(l *calendar.Calendar) string {
	if l.Solar.GetDay() <= 9 {
		return "0" + strconv.Itoa(int(l.Solar.GetDay()))
	}
	return strconv.Itoa(int(l.Solar.GetDay()))
}

func GetAllUser() []*User {
	return users
}

func readExcelFile(path string) ([]string, string, error) {
	header := make([]string, 0)
	users = users[:0]
	sheetName := ""
	f, err := excelize.OpenFile(path)
	if err != nil {
		interLogger.Error(err.Error())
		return header, sheetName, err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			interLogger.Error(err.Error())
		}
	}()

	// Get all the rows in the Sheet1.
	sheetNames := f.GetSheetList()
	if len(sheetNames) < 1 {
		interLogger.Error("sheet length error")
	}
	sheetName = sheetNames[0]
	interLogger.Debugf("fileName: %+v, sheetName: %+v", f.Path, sheetName)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		interLogger.Error(err.Error())
		return header, sheetName, err
	}

	idIdx, nameIdx, birthIdx, typeIdx := -1, -1, -1, -1
	for i, row := range rows {
		if i == 0 {
			header = row
			if idIdx, nameIdx, birthIdx, typeIdx = headerIndex(header); idIdx == -1 || nameIdx == -1 || birthIdx == -1 || typeIdx == -1 {
				msg := "header name error"
				interLogger.Errorf("%+v: %+v", msg, header)
				return header, sheetName, errors.New(msg)
			}
			interLogger.Infof("Header: %+v", header)
			continue
		}

		birthday := row[birthIdx]
		dateType := row[typeIdx]
		newBirthday := formatBirthday(f, row, sheetName, birthday, i)

		year, month, day := birthSplit(newBirthday)
		user := User{
			ID:       row[idIdx],
			Name:     row[nameIdx],
			Year:     year,
			Month:    month,
			Day:      day,
			Birthday: newBirthday,
			DateType: dateType,
		}
		interLogger.Debugf("user: %+v", user)
		users = append(users, &user)
	}
	return header, sheetName, nil
}

func birthSplit(birthday string) (int64, int64, int64) {
	sep := "-"
	year, month, day := int64(-1), int64(-1), int64(-1)
	ymd := strings.Split(birthday, sep)
	if len(ymd) >= 3 {
		i, err := strconv.Atoi(ymd[0])
		if err != nil || i < 1900 || i > 2100 {
			interLogger.Error("year value error")
		} else {
			year = int64(i)
		}
		i, err = strconv.Atoi(ymd[1])
		if err != nil || i <= 0 || i > 12 {
			interLogger.Error("year value error")
		} else {
			month = int64(i)
		}
		i, err = strconv.Atoi(ymd[2])
		if err != nil || i <= 0 || i > 31 {
			interLogger.Error("year value error")
		} else {
			day = int64(i)
		}
	}
	return year, month, day
}

func headerIndex(header []string) (int, int, int, int) {
	idIdx := indexOf(UserIDHeader, header)
	nameIdx := indexOf(UserNameHeader, header)
	birthIdx := indexOf(UserBirthHeader, header)
	dateTypeIdx := indexOf(UserDateTypeHeader, header)
	return idIdx, nameIdx, birthIdx, dateTypeIdx
}

func indexOf(element string, data []string) int {
	for k, v := range data {
		if element == v {
			return k
		}
	}
	return -1 //not found.
}

func indexOfInt(element int64, data []int64) int {
	for k, v := range data {
		if element == v {
			return k
		}
	}
	return -1 //not found.
}

func formatDate(f *excelize.File, sheetName string, cellName string) string {
	style, _ := f.NewStyle(&excelize.Style{NumFmt: 34, Lang: "ko-kr"})
	f.SetCellStyle(sheetName, cellName, cellName, style)
	e7, _ := f.GetCellValue(sheetName, cellName)
	return e7
}

func formatBirthday(f *excelize.File, row []string, sheetName, birthday string, i int) string {
	cellIdx := strconv.Itoa(i + 1)
	cellPre := toCharStr(indexOf(birthday, row) + 1)
	cellName := cellPre + cellIdx
	interLogger.Debugf("%+v, %+v, %+v, %+v th row, birthday cell: %+v", sheetName, row, birthday, i, cellName)
	return formatDate(f, sheetName, cellName)
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}

func checkResult() {

}

func writeToFile(filePath, sheetName string, newHeader []string, totYears []int64, users []*User) {
	f := excelize.NewFile() //creating a new sheet

	newSheetName := sheetName + "_result"
	idx, err := f.NewSheet(newSheetName) //creating the new sheet names
	if err != nil {
		interLogger.Errorf("%+v", err)
	}
	// set header
	for i, headerName := range newHeader {
		rowIdx := "1"
		prefix := toCharStr(i + 1)
		f.SetCellValue(newSheetName, prefix+rowIdx, headerName)
	}

	// set user
	for i, user := range users {
		rowIdx := strconv.Itoa(i + 2)
		f.SetCellValue(newSheetName, "A"+rowIdx, user.ID)
		f.SetCellValue(newSheetName, "B"+rowIdx, user.Name)
		f.SetCellValue(newSheetName, "C"+rowIdx, user.Birthday)
		f.SetCellValue(newSheetName, "D"+rowIdx, user.DateType)

		for j, year := range totYears {
			cellName := toCharStr(j+5) + rowIdx
			f.SetCellValue(newSheetName, cellName, user.FutureBirthday[year])
		}
	}

	f.SetActiveSheet(idx)
	if err := f.SaveAs(filePath); err != nil { //saving the new sheet in the file names companies
		interLogger.Errorf("%+v", err)
	}
}
