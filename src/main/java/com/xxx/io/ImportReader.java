package com.xxx.io;

import java.io.IOException;
import java.sql.SQLException;
import java.util.Map;

public interface ImportReader {
	/**
	 * 处理导入数据
	 * @throws Exception
	 */
	public void process() throws Exception;

	/**
	 * 设置读取行方式
	 * @param rowReader
	 */
	public void setRowReader(RowReader rowReader);

	/**
	 * 获取列名
	 * @return
	 */
	public Map<String, String> getHeader();

	/**
	 * 获取导入总行数
	 * @return
	 */
	public int getRowCount();

	/**
	 * 关闭
	 */
	public void close() throws IOException, SQLException;
}
