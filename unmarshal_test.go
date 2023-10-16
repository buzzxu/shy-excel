package shyexcel

import (
	"github.com/vmihailenco/msgpack/v5"
	"testing"
)

func TestMsgpack(t *testing.T) {
	table, err := File("./example.json")
	if err != nil {
		panic(err)
	}
	b, err := msgpack.Marshal(table)
	if err != nil {
		panic(err)
	}
	var _table = &Table{}
	err = msgpack.Unmarshal(b, _table)
	if err != nil {
		panic(err)
	}
}
