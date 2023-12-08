package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
)

func getAlphabetKeys() []string {
	keys := make([]string, 0)
	for i := 'A'; i <= 'Z'; i++ {
		keys = append(keys, string(i))
	}
	return keys
}

func structToInterfaceSlice(data interface{}) ([]interface{}, error) {
	dataType := reflect.TypeOf(data)
	result := make([]interface{}, 0)
	if dataType.Kind() != reflect.Struct {
		return result, errors.New(fmt.Sprintf("data 는 Struct 가 아닙니다.(dataType:%s)", dataType.Kind()))
	}
	val := reflect.Indirect(reflect.ValueOf(data))
	for k := 0; k < val.NumField(); k++ {
		result = append(result, val.Field(k).Interface())
	}
	return result, nil
}

func fileExists(fileName string) (bool, error) {
	_, err := os.Stat(fileName)
	if err == nil {
		return true, nil
	}
	return false, err
}

func saveAsExcelFile(sheetName string, fileName string, arrays [][]interface{}) (string, error) {
	rows := make([]map[string]interface{}, 0)
	rowIndex := 1
	sheetKeys := getAlphabetKeys()

	for _, array := range arrays {
		row := map[string]interface{}{}
		for k, value := range array {
			row[fmt.Sprintf("%s%v", sheetKeys[k], rowIndex)] = value
		}
		rows = append(rows, row)
		rowIndex++
	}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			panic(err)
		}
	}()

	index, _ := f.NewSheet(sheetName)
	for _, row := range rows {
		for k, v := range row {
			f.SetCellValue(sheetName, k, v)
		}
	}

	f.SetActiveSheet(index)
	if err := f.SaveAs(fileName); err != nil {
		return "", err
	}

	return fileName, nil
}
