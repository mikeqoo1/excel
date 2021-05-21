package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	var style int
	var err error
	Sheet1 := "Sheet1"
	f := excelize.NewFile()
	//創建一個工作表
	index := f.NewSheet(Sheet1)
	//調整寬度
	f.SetColWidth("Sheet1", "A", "E", 30)
	//合併A1到I1的單元格
	f.MergeCell(Sheet1, "A1", "J1")
	//設定單元格内容
	f.SetCellValue(Sheet1, "A1", "發票位置")
	//設定第一行的高度是50
	f.SetRowHeight(Sheet1, 1, 50)
	//設定單元格内容中的文字顏色大小
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":24, "color": "#777777"},
		"alignment":{"horizontal":"center", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A1", "J1", style)

	f.MergeCell(Sheet1, "A2", "C2")
	f.SetCellValue(Sheet1, "A2", "Bill to: DD")
	f.SetRowHeight(Sheet1, 2, 30)
	style, err = f.NewStyle(`{"font":{"bold":true,"size":15},"alignment":{"vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A2", "C2", style)

	f.SetCellValue(Sheet1, "A3", "time!?")
	f.SetRowHeight(Sheet1, 3, 30)
	style, err = f.NewStyle(`{"font":{"size":10},"alignment":{"vertical":"center","horizontal":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A3", "A3", style)

	//設定b這一列寬度100
	//f.SetColWidth(Sheet1, "B", "B", 100)

	f.SetCellValue(Sheet1, "A4", "開分")
	f.SetCellValue(Sheet1, "B4", "贈分")
	f.SetCellValue(Sheet1, "C4", "贈分分潤")
	f.SetCellValue(Sheet1, "D4", "洗分")
	f.SetCellValue(Sheet1, "E4", "反轉")
	f.SetCellValue(Sheet1, "A5", 111)
	f.SetCellValue(Sheet1, "B5", 111)
	f.SetCellValue(Sheet1, "C5", 111)
	f.SetCellValue(Sheet1, "D5", 111)
	f.SetCellValue(Sheet1, "E5", 111)
	f.SetCellValue(Sheet1, "A6", "GRS")
	f.SetCellValue(Sheet1, "B6", "GNSCS")
	f.SetCellValue(Sheet1, "C6", "Earnings")
	f.SetCellValue(Sheet1, "D6", "Diccount")
	f.SetCellValue(Sheet1, "E6", "LF")
	f.SetCellValue(Sheet1, "A7", 123)
	f.SetCellValue(Sheet1, "B7", 123)
	f.SetCellValue(Sheet1, "C7", 123)
	f.SetCellValue(Sheet1, "D7", 123)
	f.SetCellValue(Sheet1, "E7", 123)


	f.MergeCell(Sheet1, "A8", "A9")
	f.SetCellValue(Sheet1, "A9", "Total")



	f.SetCellValue(Sheet1, "A10", 12313)
	style, err = f.NewStyle(`{
		"font":{"bold":true, "size":40, "color": "#FF0000"},
		"alignment":{"horizontal":"center", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A10", "A10", style)

	f.MergeCell(Sheet1, "A11", "J11")
	f.SetCellValue(Sheet1, "A11", "Thank you for business.")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A11", "J11", style)


	f.MergeCell(Sheet1, "A12", "J12")
	f.SetCellValue(Sheet1, "A12", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A12", "J12", style)

	f.MergeCell(Sheet1, "A13", "J13")
	f.SetCellValue(Sheet1, "A13", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A13", "J13", style)

	f.MergeCell(Sheet1, "A14", "J14")
	f.SetCellValue(Sheet1, "A14", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A14", "J14", style)

	f.MergeCell(Sheet1, "A15", "J15")
	f.SetCellValue(Sheet1, "A15", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A15", "J15", style)

	f.MergeCell(Sheet1, "A16", "J16")
	f.SetCellValue(Sheet1, "A16", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A16", "J16", style)

	f.MergeCell(Sheet1, "A17", "J17")
	f.SetCellValue(Sheet1, "A17", "公式")
	style, err = f.NewStyle(`{
		"fill":{"type":"pattern", "color":["#E0EBF5"], "pattern":1},
		"font":{"bold":true, "size":10, "color": "#777777"},
		"alignment":{"horizontal":"right", "vertical":"center"}}`)
	if err != nil {
		fmt.Println(err)
	}
	f.SetCellStyle(Sheet1, "A17", "J17", style)

	//下面是細節...
	f.MergeCell(Sheet1, "A18", "I18")
	f.SetCellValue(Sheet1, "A18", "細節時間")

	f.SetCellValue(Sheet1, "A19", "開分")
	f.SetCellValue(Sheet1, "B19", "店特殊贈分")
	f.SetCellValue(Sheet1, "C19", "日首儲贈分攤")
	f.SetCellValue(Sheet1, "D19", "洗分")
	f.SetCellValue(Sheet1, "E19", "取消上一筆")
	f.SetCellValue(Sheet1, "F19", "純淨利")
	f.SetCellValue(Sheet1, "G19", "上層分潤")
	f.SetCellValue(Sheet1, "H19", "折扣")
	f.SetCellValue(Sheet1, "I19", "總結分潤")

	//設定活頁簿的默認工作表
	f.SetActiveSheet(index)
	//根據指定路徑儲存檔案
	err = f.SaveAs("invoice.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
