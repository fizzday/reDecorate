package gooffice

import (
	"github.com/tealeg/xlsx"
	"fmt"
	"errors"
	"github.com/gohouse/gorose/utils"
)

// ExportExcel 导出表格
// filePath 文件路径,如: static/export.xlsx
// datas 所有数据,如:
//	var datas = []map[string]interface{}{
//		{"a": 1, "b": 2, "c": 3},
//		{"a": 11, "b": 12, "c": 13},
//		{"a": "搞事情", "b": "UTF-8", "c": "汉字"},
//	}
// tableHead 表头, 如:
//	var h = map[string]string{
//		"id":       "",
//		"username": "用户名",
//		"age":      "年龄",
//	}
type Excel struct {
	tableHead map[string]interface{}
	filePath  string
	Sheet     []string
}

func NewExcel() *Excel {
	return &Excel{}
}

func (e *Excel) TableHead(th map[string]interface{}) *Excel {
	e.tableHead = th
	return e
}

func (e *Excel) FilePath(fp string) *Excel {
	e.filePath = fp
	return e
}

func (e *Excel) ExportExcel(datas []map[string]interface{}) error {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
		return err
	}
	if len(datas) == 0 {
		return errors.New("数据为空")
	}

	// 添加表头
	var keys []string
	row = sheet.AddRow()
	if len(e.tableHead) == 0 {
		for key, _ := range datas[0] {
			keys = append(keys, key)
			// 写入表头
			cell = row.AddCell()
			cell.Value = key
		}
	} else {
		//var keys2 = utils.ArrayValues(e.tableHead)
		// 写入表头
		for k, v := range e.tableHead {
			keys = append(keys, k)
			cell = row.AddCell()
			cell.Value = v.(string)
		}
	}

	// 循环写入内容
	for _, item := range datas {
		row = sheet.AddRow()
		for _, v := range keys {
			cell = row.AddCell()
			cell.Value = utils.ParseStr(item[v])

			err := file.Save(e.filePath)
			if err != nil {
				fmt.Printf(err.Error())
				return err
			}
		}
	}

	return nil
}
