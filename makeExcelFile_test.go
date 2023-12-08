package main

import (
	"github.com/stretchr/testify/assert"
	"os"
	"testing"
)

const (
	SaveFileName = "test.xlsx"
	SheetName    = "Sheet1"
)

func TestSaveAsExcelFile_interfaceSlice_배열을_엑셀파일로_저장(t *testing.T) {

	rows := make([][]interface{}, 0)
	rows = append(rows, []interface{}{"번호", "이름", "확인여부"})
	rows = append(rows, []interface{}{1, "지현", true})
	rows = append(rows, []interface{}{2, "하늘", false})
	member := struct {
		Id    int64
		Name  string
		Check bool
	}{
		Id:    3,
		Name:  "찬영",
		Check: true,
	}
	rows = append(rows, []interface{}{member.Id, member.Name, member.Check})

	fileName, _ := saveAsExcelFile(SheetName, SaveFileName, rows)
	defer func() {
		if err := os.Remove(fileName); err != nil {
			t.Error(err.Error())
		}
	}()

	exists, err := fileExists(fileName)
	if err != nil {
		t.Error(err.Error())
	}

	assert.Equal(t, exists, true)
}

func TestSaveAsExcelFile_struct_interface_배열로_바꾸고_엑셀파일로_저장(t *testing.T) {
	type Member struct {
		Id    int64
		Name  string
		Check bool
	}
	rows := make([][]interface{}, 0)
	rows = append(rows, []interface{}{"번호", "이름", "확인여부"})
	got1, _ := structToInterfaceSlice(Member{int64(1), "ann", true})
	got2, _ := structToInterfaceSlice(struct {
		Id    int64
		Name  string
		Check bool
	}{int64(2), "", false})

	rows = append(rows, got1)
	rows = append(rows, got2)

	fileName, _ := saveAsExcelFile(SheetName, SaveFileName, rows)
	defer func() {
		if err := os.Remove(fileName); err != nil {
			t.Error(err.Error())
		}
	}()

	exists, err := fileExists(fileName)
	if err != nil {
		t.Error(err.Error())
	}
	assert.Equal(t, exists, true)
}

func TestSaveAsExcelFile_파일_저장_실패(t *testing.T) {

	rows := make([][]interface{}, 0)
	rows = append(rows, []interface{}{"번호", "이름", "확인여부"})
	rows = append(rows, []interface{}{1, "지현", true})
	rows = append(rows, []interface{}{2, "하늘", false})

	fileName, err := saveAsExcelFile(SheetName, "test.unsupportedFormat", rows)
	if err != nil {
		assert.Equal(t, "unsupported workbook file format", err.Error())
	}
	defer func() {
		if err = os.Remove(fileName); err != nil {
			assert.Equal(t, "remove : no such file or directory", err.Error())
		}
	}()

}

func Test_fileExists(t *testing.T) {
	t.Run("파일이 없는 경우", func(t *testing.T) {
		exists, err := fileExists(SaveFileName)
		assert.Equal(t, false, exists)
		if err != nil {
			assert.Equal(t, true, os.IsNotExist(err))
		}
	})
}

func Test_structToInterfaceSlice(t *testing.T) {
	t.Run("struct 타입이 아닌 경우", func(t *testing.T) {
		data := []int{1, 2, 3}
		got, err := structToInterfaceSlice(data)
		assert.Equal(t, []interface{}{}, got)
		assert.Equal(t, "data 는 Struct 가 아닙니다.(dataType:slice)", err.Error())
	})

	t.Run("struct 타입이 아닌 경우", func(t *testing.T) {
		data := 111
		got, err := structToInterfaceSlice(data)
		assert.Equal(t, []interface{}{}, got)
		assert.Equal(t, "data 는 Struct 가 아닙니다.(dataType:int)", err.Error())
	})

	t.Run("struct 타입이 아닌 경우", func(t *testing.T) {
		data := "zzzzz"
		got, err := structToInterfaceSlice(data)
		assert.Equal(t, []interface{}{}, got)
		assert.Equal(t, "data 는 Struct 가 아닙니다.(dataType:string)", err.Error())
	})
}
