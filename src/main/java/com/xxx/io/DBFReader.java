package com.xxx.io;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * DBF文件Reader
 */
public class DBFReader implements ImportReader {

	private static final Logger logger = LoggerFactory.getLogger(DBFReader.class);
	/**
	 * 加载DBF驱动
	 */
	static {
		try {
			Class.forName("com.hxtt.sql.dbf.DBFDriver");
		} catch (ClassNotFoundException e) {
			logger.warn("加载DBF驱动异常：" + e);
		}
	}

	private Connection conn = null;
	private Statement stm = null;
	private ResultSet resultSet = null;

	//存储行记录的容器
	private Map<String, String> rowData = new HashMap<String, String>();
	private Map<String, String> header = new LinkedHashMap<String, String>();

	private RowReader rowReader;
	private int rowCount = 0;
	private int columnCount = 0;

	/**
	 * dbf文件路径（包含后缀名）
	 * @param path
	 * @throws SQLException 
	 */
	public DBFReader(String path) throws SQLException {
		//在windows下替换“\\”为/
		path = path.replace("\\", "/");
		int endIndex = path.lastIndexOf("/") + 1;
		String DBUrl = path.substring(0, endIndex);
		String tableName = path.substring(endIndex, path.lastIndexOf("."));
		String url = "jdbc:dbf:/" + DBUrl + "?charSet=GBK";
		conn = DriverManager.getConnection(url);
		stm = conn.createStatement();
		resultSet = stm.executeQuery("SELECT COUNT(*) FROM " + tableName);
		if (resultSet.next()) {
			rowCount = resultSet.getInt(1);
		}
		resultSet = stm.executeQuery("SELECT * FROM " + tableName);
		ResultSetMetaData rsmd = resultSet.getMetaData();
		columnCount = rsmd.getColumnCount();
		for (int i = 1; i <= columnCount; i++) {
			header.put(String.valueOf(i), rsmd.getColumnName(i));
		}
	}

	/**
	 * 设置读取行方式
	 * @param rowReader
	 */
	@Override
	public void setRowReader(RowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**
	 * 获取列名
	 * @return
	 */
	@Override
	public Map<String, String> getHeader() {
		return header;
	}

	/**
	 * 处理导入数据
	 */
	@Override
	public void process() throws SQLException {
		if (rowReader != null) {
			while (resultSet.next()) {
				for (int i = 1; i <= columnCount; i++) {
					String value = resultSet.getString(i);
					if (value == null) {
						value = "";
					}
					rowData.put(String.valueOf(i), value);
				}
				rowReader.dealRowData(resultSet.getRow(), rowData);
				rowData.clear();
			}
		}
	}

	/**
	 * 关闭连接
	 * @throws SQLException 
	 */
	@Override
	public void close() throws SQLException {
		if (resultSet != null) {
			resultSet.close();
		}
		if (stm != null) {
			stm.close();
		}
		if (conn != null) {
			conn.close();
		}
	}

	/**
	 * 获取导入总行数
	 * @return
	 */
	@Override
	public int getRowCount() {
		return rowCount;
	}
}
