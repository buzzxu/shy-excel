package shyexcel

import (
	"context"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
)

func File(file string) (*Table, error) {
	bytes, err := os.ReadFile(file)
	if err != nil {
		return nil, err
	}
	return Json(bytes)
}

func Http(url, method string, responseType ResponseType, funcHeader func(header http.Header)) (*Table, error) {
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
		defer resp.Body.Close()
	}
	if resp.StatusCode == 200 {
		b, err := io.ReadAll(resp.Body)
		if err != nil {
			return nil, err
		}
		switch responseType {
		case JSON:
			return Json(b)
		case MsgPack:
			return Msgpack(b)
		default:
			return Json(b)
		}
	}
	return nil, errors.New(fmt.Sprintf("http resp code: %d", resp.StatusCode))
}
