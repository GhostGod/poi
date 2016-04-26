package com.xxx.io;

import java.util.Map;

public interface RowReader {

	/**
	 * 处理一行数据，实现业务逻辑
	 * @param curRow 当前行
	 * @param rowData 当前行的数据
	 */
	public void dealRowData(int curRow, Map<String, String> rowData);
}
