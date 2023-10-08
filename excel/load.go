package excel

import (
	"context"
	"errors"
	"fmt"
	jsoniter "github.com/json-iterator/go"
	"io"
	"net/http"
	"os"
)

func File(file string) (*Table, error) {
	bytes, err := os.ReadFile(file)
	if err != nil {
		return nil, err
	}
	var table = &Table{}
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	err = json.Unmarshal(bytes, table)
	if err != nil {
		return nil, err
	}
	return table, nil
}

func JSON(str string) (*Table, error) {
	var table = &Table{}
	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	err := json.Unmarshal([]byte(str), table)
	if err != nil {
		return nil, err
	}
	return table, nil
}

func HTTP(url, method string, funcHeader func(header http.Header)) (*Table, error) {
	req, err := http.NewRequestWithContext(context.Background(), method, url, nil)
	if err != nil {
		return nil, err
	}
	if funcHeader != nil {
		funcHeader(req.Header)
	}
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, err
	}
	if resp != nil && resp.Body != nil {
		resp.Body.Close()
	}
	if resp.StatusCode == 200 {
		b, err := io.ReadAll(resp.Body)
		if err != nil {
			return nil, err
		}
		var json = jsoniter.ConfigCompatibleWithStandardLibrary
		var table = &Table{}
		err = json.Unmarshal(b, table)
		if err != nil {
			return nil, err
		}
		return table, nil
	}
	return nil, errors.New(fmt.Sprintf("http resp code: %d", resp.StatusCode))
}
