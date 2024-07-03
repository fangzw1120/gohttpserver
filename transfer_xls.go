package main

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/Lofanmi/chinese-calendar-golang/calendar"
	interLogger "github.com/codeskyblue/gohttpserver/logger"
	"github.com/xuri/excelize/v2"
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
	Extra          string
}

const (
	UserIDHeader       = "工号"
	UserNameHeader     = "姓名"
	UserBirthHeader    = "生日"
	UserDateTypeHeader = "日历类型"
)

func TransferInit() {
	basePath := GetMainDirectory()
	logPath := basePath + "logs/"

	interLogger.Init(logPath, "transfer_xls.log", false, false, true)
	interLogger.Infof("config: %+v", gcfg)
}

// GetMainDirectory 获取进程所在目录: 末尾带反斜杠
func GetMainDirectory() string {
	path, err := filepath.Abs(os.Args[0])

	if err != nil {
		return ""
	}

	fullPath := filepath.Dir(path)
	return pathAddBackslash(fullPath)
}

// PathAddBackslash 路径最后添加反斜杠
func pathAddBackslash(path string) string {
	i := len(path) - 1

	if !os.IsPathSeparator(path[i]) {
		path += string(os.PathSeparator)
	}
	return path
}

func generateOutFileName(fileName string) string {
	extension := filepath.Ext(fileName)                 // 获取文件后缀名
	baseName := strings.TrimSuffix(fileName, extension) // 去除后缀名的文件名

	//// 查找第一个点符号的索引
	//dotIndex := strings.LastIndex(baseName, ".")
	//
	//// 如果找到点符号，则截取点符号之前的部分作为真正的文件名
	//if dotIndex != -1 {
	//	baseName = baseName[:dotIndex]
	//}

	// 拼接真正的文件名和 "_res"，再加上原来的后缀名
	realFileName := baseName + "_res" + extension
	return realFileName
}

func transferXLS(filePath, outputPath string) {
	// parse xls, get birthday and type
	header, sheetName, err := readExcelFile(filePath)
	if err != nil {
		interLogger.Errorf("%+v", err)
		return
	}
	interLogger.Infof("sheetName: %+v, header %+v", sheetName, header)

	// get parse result
	users = GetAllUser()

	// target years list
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

	// generate target year birthday
	cnt := 0
	for i, user := range users {
		cnt++
		//logger.Infof("%+v: user: %+v", i+1, user)
		futureBirthday := make(map[int64]string, totYears)
		extra := ""
		for _, year := range yearLT {
			if user.DateType == "农历" || user.DateType == "阴历" {
				// 目标年份的农历时间对象
				isLeapMonth := false
				// valLunarItem := calendar.ByLunar(year, user.Month, user.Day, 0, 0, 0, isLeapMonth)
				ts := ToSolarTimestamp(year, user.Month, user.Day, 0, 0, 0, isLeapMonth)
				valLunarItem := calendar.ByTimestamp(ts)

				maxDays := lunarDays(year, user.Month)
				if isLeapMonth {
					maxDays = leapDays(year)
				}
				oldDay := user.Day
				newDay := ReCorrectDay(oldDay, maxDays)
				if oldDay != newDay {
					extra += fmt.Sprintf("农历%+v年-%+v月-%+v日, 转为%+v月-%+v日;", year, user.Month, oldDay, user.Month, newDay)
				}
				// 对应年份公历
				valSolar := valLunarItem.Solar
				// 目标年份农历对应的阳历日期
				futureBirthday[year] = fmt.Sprintf("%+v-%+v-%+v", valSolar.GetYear(), GetSolarMonthStr(valLunarItem), GetSolarDayStr(valLunarItem))
				//// 对应年份农历
				//valLunar := valLunarItem.Lunar
				//logger.Debugf("年份: %+v, 农历: %+v-%+v-%+v, Solar: %+v-%+v-%+v", year,
				//	valLunar.GetYear(), valLunar.GetMonth(), valLunar.GetDay(),
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
		user.Extra = extra
		interLogger.Infof("%+v: user: %+v", i+1, user)
	}
	interLogger.Infof("tot user: %+v", cnt)

	// new header and output
	newHeader := append(header, yearStrLt...)
	writeToFile(outputPath, sheetName, newHeader, yearLT, users)
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

func readExcelFile(path string) ([]string, string, error) {
	header := make([]string, 0)
	users = users[:0]
	sheetName := ""

	// open file
	f, err := excelize.OpenFile(path)
	if err != nil {
		interLogger.Error(err.Error())
		return header, sheetName, err
	}
	defer func() {
		// Close the spreadsheet.
		if err = f.Close(); err != nil {
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

	// loop row
	idIdx, nameIdx, birthIdx, typeIdx := -1, -1, -1, -1
	for i, row := range rows {
		// first row as header
		if i == 0 {
			header = row
			if idIdx, nameIdx, birthIdx, typeIdx = headerIndex(header); idIdx == -1 || nameIdx == -1 || birthIdx == -1 || typeIdx == -1 {
				msg := "header name error"
				interLogger.Errorf("%+v: %+v", msg, header)
				return header, sheetName, errors.New(msg)
			}
			interLogger.Debugf("Header: %+v", header)
			continue
		}

		// birthday column to correct format
		birthday := row[birthIdx]
		dateType := row[typeIdx]
		newBirthday := formatBirthday(f, row, sheetName, birthday, i)
		interLogger.Debugf("%+v: %+v, newBirthday %+v, dateType %+v", i, row, newBirthday, dateType)

		// birthday to year, month, day
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
		interLogger.Debugf("%+v: user: %+v", i, user)
		users = append(users, &user)
	}
	return header, sheetName, nil
}

func GetAllUser() []*User {
	return users
}

// birthSplit ...
// @Description: parse birthday
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

// formatBirthday ...
// @Description: birthday to correct format
func formatBirthday(f *excelize.File, row []string, sheetName, birthday string, i int) string {
	cellIdx := strconv.Itoa(i + 1)
	cellPre := toCharStr(indexOf(birthday, row) + 1)
	cellName := cellPre + cellIdx
	return formatDate(f, sheetName, cellName)
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}

func writeToFile(filePath, sheetName string, newHeader []string, totYears []int64, users []*User) {
	f := excelize.NewFile() //creating a new sheet

	newSheetName := sheetName + "_result"
	idx, err := f.NewSheet(newSheetName) //creating the new sheet names
	if err != nil {
		interLogger.Errorf("%+v", err)
		return
	}
	// set header
	for i, headerName := range newHeader {
		rowIdx := "1"
		prefix := toCharStr(i + 1)
		f.SetCellValue(newSheetName, prefix+rowIdx, headerName)
		interLogger.Debugf("%+v: %+v, %+v", newSheetName, prefix+rowIdx, headerName)
	}

	// set user
	for i, user := range users {
		rowIdx := strconv.Itoa(i + 2)
		f.SetCellValue(newSheetName, "A"+rowIdx, user.ID)
		f.SetCellValue(newSheetName, "B"+rowIdx, user.Name)
		f.SetCellValue(newSheetName, "C"+rowIdx, user.Birthday)
		f.SetCellValue(newSheetName, "D"+rowIdx, user.DateType)
		//logger.Debugf("%+v: %+v, %+v", newSheetName, rowIdx, user)

		for j, year := range totYears {
			cellName := toCharStr(j+5) + rowIdx
			f.SetCellValue(newSheetName, cellName, user.FutureBirthday[year])
			//logger.Debugf("%+v: %+v, %+v", newSheetName, cellName, user.FutureBirthday[year])
		}
		cellName := toCharStr(len(totYears)+5) + rowIdx
		f.SetCellValue(newSheetName, cellName, user.Extra)
	}

	f.SetActiveSheet(idx)
	if err = f.SaveAs(filePath); err != nil { //saving the new sheet in the file names companies
		interLogger.Errorf("%+v", err)
		return
	}
}

var lunars = [...]int64{
	0x04bd8, 0x04ae0, 0x0a570, 0x054d5, 0x0d260, 0x0d950, 0x16554, 0x056a0, 0x09ad0, 0x055d2, // 1900-1909
	0x04ae0, 0x0a5b6, 0x0a4d0, 0x0d250, 0x1d255, 0x0b540, 0x0d6a0, 0x0ada2, 0x095b0, 0x14977, // 1910-1919
	0x04970, 0x0a4b0, 0x0b4b5, 0x06a50, 0x06d40, 0x1ab54, 0x02b60, 0x09570, 0x052f2, 0x04970, // 1920-1929
	0x06566, 0x0d4a0, 0x0ea50, 0x06e95, 0x05ad0, 0x02b60, 0x186e3, 0x092e0, 0x1c8d7, 0x0c950, // 1930-1939
	0x0d4a0, 0x1d8a6, 0x0b550, 0x056a0, 0x1a5b4, 0x025d0, 0x092d0, 0x0d2b2, 0x0a950, 0x0b557, // 1940-1949
	0x06ca0, 0x0b550, 0x15355, 0x04da0, 0x0a5b0, 0x14573, 0x052b0, 0x0a9a8, 0x0e950, 0x06aa0, // 1950-1959
	0x0aea6, 0x0ab50, 0x04b60, 0x0aae4, 0x0a570, 0x05260, 0x0f263, 0x0d950, 0x05b57, 0x056a0, // 1960-1969
	0x096d0, 0x04dd5, 0x04ad0, 0x0a4d0, 0x0d4d4, 0x0d250, 0x0d558, 0x0b540, 0x0b6a0, 0x195a6, // 1970-1979
	0x095b0, 0x049b0, 0x0a974, 0x0a4b0, 0x0b27a, 0x06a50, 0x06d40, 0x0af46, 0x0ab60, 0x09570, // 1980-1989
	0x04af5, 0x04970, 0x064b0, 0x074a3, 0x0ea50, 0x06b58, 0x055c0, 0x0ab60, 0x096d5, 0x092e0, // 1990-1999
	0x0c960, 0x0d954, 0x0d4a0, 0x0da50, 0x07552, 0x056a0, 0x0abb7, 0x025d0, 0x092d0, 0x0cab5, // 2000-2009
	0x0a950, 0x0b4a0, 0x0baa4, 0x0ad50, 0x055d9, 0x04ba0, 0x0a5b0, 0x15176, 0x052b0, 0x0a930, // 2010-2019
	0x07954, 0x06aa0, 0x0ad50, 0x05b52, 0x04b60, 0x0a6e6, 0x0a4e0, 0x0d260, 0x0ea65, 0x0d530, // 2020-2029
	0x05aa0, 0x076a3, 0x096d0, 0x04afb, 0x04ad0, 0x0a4d0, 0x1d0b6, 0x0d250, 0x0d520, 0x0dd45, // 2030-2039
	0x0b5a0, 0x056d0, 0x055b2, 0x049b0, 0x0a577, 0x0a4b0, 0x0aa50, 0x1b255, 0x06d20, 0x0ada0, // 2040-2049
	0x14b63, 0x09370, 0x049f8, 0x04970, 0x064b0, 0x168a6, 0x0ea50, 0x06b20, 0x1a6c4, 0x0aae0, // 2050-2059
	0x0a2e0, 0x0d2e3, 0x0c960, 0x0d557, 0x0d4a0, 0x0da50, 0x05d55, 0x056a0, 0x0a6d0, 0x055d4, // 2060-2069
	0x052d0, 0x0a9b8, 0x0a950, 0x0b4a0, 0x0b6a6, 0x0ad50, 0x055a0, 0x0aba4, 0x0a5b0, 0x052b0, // 2070-2079
	0x0b273, 0x06930, 0x07337, 0x06aa0, 0x0ad50, 0x14b55, 0x04b60, 0x0a570, 0x054e4, 0x0d160, // 2080-2089
	0x0e968, 0x0d520, 0x0daa0, 0x16aa6, 0x056d0, 0x04ae0, 0x0a9d4, 0x0a2d0, 0x0d150, 0x0f252, // 2090-2099
	0x0d520, // 2100
}

func leapMonth(year int64) int64 {
	return lunars[year-1900] & 0xf
}

func leapDays(year int64) (days int64) {
	if leapMonth(year) == 0 {
		days = 0
	} else if (lunars[year-1900] & 0x10000) != 0 {
		days = 30
	} else {
		days = 29
	}
	return
}

func lunarDays(year, month int64) (days int64) {
	if month > 12 || month < 1 {
		days = 0
	} else if (lunars[year-1900] & (0x10000 >> uint64(month))) != 0 {
		days = 30
	} else {
		days = 29
	}
	return
}

func daysOfLunarYear(year int64) int64 {
	var (
		i, sum int64
	)
	sum = 29 * 12
	for i = 0x8000; i > 0x8; i >>= 1 {
		if (lunars[year-1900] & i) != 0 {
			sum++
		}
	}
	return sum + leapDays(year)
}

func ReCorrectDay(day, maxDays int64) int64 {
	if day <= maxDays {
		return day
	}
	// 出生农历有的日期在目标年没有
	// 例如不是每一年都有大年30，那就用大年29代替

	target := day
	for {
		if target <= 1 {
			return 1
		}
		if target <= maxDays {
			return target
		}
		target -= 1
	}
}

// ToSolarTimestamp 转换为时间戳
func ToSolarTimestamp(year, month, day, hour, minute, second int64, isLeapMonth bool) int64 {
	var (
		i, offset int64
	)
	// 参数合法性效验
	if year < 1900 || year > 2100 {
		return 0
	}
	// 参数区间 1900.1.31~2100.12.1
	m := leapMonth(year)
	// 传参要求计算该闰月公历 但该年得出的闰月与传参的月份并不同
	if isLeapMonth && (m != month) {
		isLeapMonth = false
	}
	// 超出了最大极限值
	if 2100 == year && 12 == month && day > 1 || 1900 == year && 1 == month && day < 31 {
		return 0
	}
	days := lunarDays(year, month)
	maxDays := days
	// if month is leap, _day use leapDays method
	if isLeapMonth {
		maxDays = leapDays(year)
	}
	// 参数合法性效验
	// if day > maxDays {
	// 	return 0
	// }
	day = ReCorrectDay(day, maxDays)

	// 计算农历的时间差
	offset = 0
	for i = 1900; i < year; i++ {
		offset += daysOfLunarYear(i)
	}
	isAdd := false
	for i = 1; i < month; i++ {
		leap := leapMonth(year)
		if !isAdd {
			// 处理闰月
			if leap <= i && leap > 0 {
				offset += leapDays(year)
				isAdd = true
			}
		}
		offset += lunarDays(year, i)
	}
	// 转换闰月农历 需补充该年闰月的前一个月的时差
	if isLeapMonth {
		offset += days
	}
	// 1900 年农历正月初一的公历时间为 1900年1月30日0时0分0秒 (该时间也是本农历的最开始起始点)
	// startTimestamp := time.Date(1900, 1, 30, 0, 0, 0, 0, time.Local).Unix()
	var startTimestamp int64 = -2206512000

	return (offset+day)*86400 + startTimestamp + hour*3600 + minute*60 + second
}
