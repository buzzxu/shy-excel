package shyexcel

import (
	jsoniter "github.com/json-iterator/go"
	"github.com/vmihailenco/msgpack/v5"
)

func Json(bytes []byte) (*Table, error) {
	var table = &Table{}
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	err := json.Unmarshal(bytes, table)
	if err != nil {
		return nil, err
	}
	return table, nil
}

func Msgpack(bytes []byte) (*Table, error) {
	var table = &Table{}
	err := msgpack.Unmarshal(bytes, table)
	if err != nil {
		return nil, err
	}
	return table, nil
}
