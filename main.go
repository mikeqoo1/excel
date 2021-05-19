package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	f := excelize.NewFile()
	// 創建一個工作表
	index := f.NewSheet("Sheet1")
	// 設定儲存格的值
	f.SetCellValue("Sheet1", "A1", "發票位置大小!?!?")
	f.SetCellValue("Sheet1", "A2", "Bill to: DD")
	f.SetCellValue("Sheet1", "A3", "time!?")
	f.SetCellValue("Sheet1", "A4", "開分")
	f.SetCellValue("Sheet1", "B4", "贈分")
	f.SetCellValue("Sheet1", "C4", "贈分分潤")
	f.SetCellValue("Sheet1", "D4", "洗分")
	f.SetCellValue("Sheet1", "E4", "反轉")
	f.SetCellValue("Sheet1", "A5", "開分數值")
	f.SetCellValue("Sheet1", "B5", "贈分數值")
	f.SetCellValue("Sheet1", "C5", "贈分分潤數值")
	f.SetCellValue("Sheet1", "D5", "洗分數值")
	f.SetCellValue("Sheet1", "E5", "反轉數值")
	f.SetCellValue("Sheet1", "A6", "GRS")
	f.SetCellValue("Sheet1", "B6", "GNSCS")
	f.SetCellValue("Sheet1", "C6", "Earnings")
	f.SetCellValue("Sheet1", "D6", "Diccount")
	f.SetCellValue("Sheet1", "E6", "LF")
	f.SetCellValue("Sheet1", "A7", "GRS數值")
	f.SetCellValue("Sheet1", "B7", "GNSCS數值")
	f.SetCellValue("Sheet1", "C7", "Earnings數值")
	f.SetCellValue("Sheet1", "D7", "Diccount數值")
	f.SetCellValue("Sheet1", "E7", "LF數值")
	f.SetCellValue("Sheet1", "A9", "Total")
	f.SetCellValue("Sheet1", "A10", "Total數值")
	f.SetCellValue("Sheet1", "A11", "          Thank you for business.")
	f.SetCellValue("Sheet1", "A12", "公式1")
	f.SetCellValue("Sheet1", "A13", "公式2")
	f.SetCellValue("Sheet1", "A14", "公式3")
	f.SetCellValue("Sheet1", "A15", "公式4")
	f.SetCellValue("Sheet1", "A16", "公式5")
	f.SetCellValue("Sheet1", "A17", "公式6")
	//下面是細節...
	f.SetCellValue("Sheet1", "A18", "細節時間")

	// 設定活頁簿的默認工作表
	f.SetActiveSheet(index)
	// 根據指定路徑儲存檔案
	err := f.SaveAs("發票.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
