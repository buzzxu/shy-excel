package shyexcel

// calcMaxRow 计算给定数据结构的最大深度
func calcMaxRow(columns []*Column, data map[string]interface{}) int {
	var maxDepth = 0
	for _, col := range columns {
		if col.Collection {
			if nestedData, ok := data[col.Name]; ok {
				list, ok := nestedData.([]interface{})
				if !ok {
					continue
				}
				maxDepth = len(list)
				var depth = 0
				for _, item := range list {
					nestedData := item.(map[string]interface{})
					current_row := calcMaxRow(col.Columns, nestedData)
					if current_row == 0 {
						nestedData["__rows__"] = current_row + 1
					} else {
						nestedData["__rows__"] = current_row
					}

					depth = depth + current_row
				}
				if depth >= maxDepth {
					return depth
				}
				break
			}
			break
		}
	}
	return maxDepth
}
