package com.xxx.io;

import java.io.IOException;
import java.sql.SQLException;
import java.util.List;

/**
 * 导出Writer接口
 * @author liuyang
 * @date 2014年10月30日 下午3:58:17
 */
public interface ExportWriter {

	/**
	 * 初始化标题列
	 */
	public void initHeader() throws SQLException, IOException;

	/**
	 * 追加数据
	 * @param dataList 数据list
	 */
	public void appendData(List<?> dataList) throws Exception;

	/**
	 * 关闭输出流，保存文件
	 */
	public void close() throws IOException, SQLException;

}
