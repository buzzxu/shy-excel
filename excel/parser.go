package excel

import (
	jsoniter "github.com/json-iterator/go"
	"os"
)

func fromFile(file string) *Table {
	bytes, err := os.ReadFile(file)
	if err != nil {
		panic(err)
	}
	return jsonParser(bytes)
}

func jsonParser(bytes []byte) *Table {
	var excel = &Table{}
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	json.Unmarshal(bytes, excel)
	return excel
}
