package com.xxx.io;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.xxx.contant.Constant;
import com.xxx.io.ExportField.DataType;

/**
 * DBF文件Writer
 * @author liuyang
 * @date 2014年11月2日 上午9:23:32
 */
public class DBFWriter extends AbstractExportWriter {

	private static final Logger logger = LoggerFactory.getLogger(DBFWriter.class);

	//加载驱动
	static {
		try {
			Class.forName("com.hxtt.sql.dbf.DBFDriver");
		} catch (ClassNotFoundException e) {
			logger.warn("加载DBF驱动异常：" + e);
		}
	}

	private Connection con = null;
	private PreparedStatement stmt = null;
	private String dataBasePath;
	private String tableName;

	private int columns;
	private int rows;

	/**
	 * 构造方法
	 * @param filePath 文件路径
	 * @param fields 字段名列表
	 * @throws SQLException 
	 */
	public DBFWriter(String filePath, List<ExportField> fields) throws SQLException {
		int index = filePath.lastIndexOf(Constant.SEPARATOR_SLASH);
		this.dataBasePath = filePath.substring(0, index);
		this.tableName = filePath.substring(index + 1).split(Constant.SEPARATOR_BACKSLASH + Constant.SEPARATOR_DOT)[0];
		this.fields = fields;
		this.columns = fields.size();
		String url = "jdbc:dbf:///" + this.dataBasePath + "?charSet=GBK";
		con = DriverManager.getConnection(url);
		con.setAutoCommit(false);
	}

	/**
	 * 构造方法
	 * @param dataBasePath 数据库地址
	 * @param tableName 表名
	 * @param fields 字段名列表
	 * @throws SQLException 
	 */
	public DBFWriter(String dataBasePath, String tableName, List<ExportField> fields) throws SQLException {
		this.dataBasePath = dataBasePath;
		this.tableName = tableName;
		this.fields = fields;
		this.columns = fields.size();
		String url = "jdbc:dbf:///" + this.dataBasePath + "?charSet=GBK";
		con = DriverManager.getConnection(url);
		con.setAutoCommit(false);
	}

	/**
	 * 创建表
	 * @throws SQLException 
	 */
	@Override
	public void initHeader() throws SQLException {
		StringBuffer createTableSQL = new StringBuffer("CREATE TABLE ");
		createTableSQL.append(tableName).append("(");
		for (int i = 0; i < columns; i++) {
			String fieldName = fields.get(i).getShowName();
			if (fieldName.length() > 10) {
				fieldName = fieldName.substring(0, 10);
			}
			createTableSQL.append(fieldName.toUpperCase()).append(" ");
			DataType dataType = fields.get(i).getType();
			switch (dataType) {
			case STRING:
				createTableSQL.append("VARCHAR(").append(fields.get(i).getLength()).append(")");
				break;
			case INT:
				createTableSQL.append("INT");
				break;
			case DATE:
				createTableSQL.append("DATE");
				break;
			case BOOLEAN:
				createTableSQL.append("BOOLEAN");
				break;
			default:
				createTableSQL.append("VARCHAR(").append(fields.get(i).getLength()).append(")");
				break;
			}
			createTableSQL.append(Constant.SEPARATOR_COMMA);
		}
		if (hasMessage) {
			createTableSQL.append(MESSAGE).append(Constant.SEPARATOR_COMMA);
		}
		String createSQL = createTableSQL.substring(0, createTableSQL.length() - 1) + ");";
		stmt = con.prepareStatement(createSQL);
		stmt.executeUpdate();
		con.commit();
		stmt.close();
	}

	@Override
	public void appendData(List<?> dataList) throws Exception {
		for (Object o : dataList) {
			appendData(o, null);
		}
	}

	@Override
	public void close() throws SQLException {
		if (stmt != null) {
			stmt.close();
		}
		if (con != null) {
			con.close();
		}
	}

	/**
	 * 追加一条数据
	 */
	private void appendData(Object data, String message) throws Exception {
		Object[] objs = parseData(data, message);
		StringBuffer insertDataSQL = new StringBuffer("INSERT INTO " + tableName + " VALUES(?");
		for (int i = 0, size = columns - 1; i < size; i++) {
			insertDataSQL.append(",?");
		}
		insertDataSQL.append(");");
		stmt = con.prepareStatement(insertDataSQL.toString());
		for (int i = 0, size = columns; i < size; i++) {
			stmt.setObject(i + 1, objs[i]);
		}
		stmt.executeUpdate();
		con.commit();
		stmt.close();
		rows++;
	}

	public int getColumns() {
		return columns;
	}

	public int getRows() {
		return rows;
	}

	/**
	 * 删除表方法
	 * @param filePath DBF文件路径
	 * @return
	 */
	public static boolean dropTable(String filePath) {
		int index = filePath.replace(Constant.SEPARATOR_BACKSLASH, Constant.SEPARATOR_SLASH).lastIndexOf(
				Constant.SEPARATOR_SLASH);
		String dataBasePath = filePath.substring(0, index);
		String tableName = filePath.substring(index + 1).split(Constant.SEPARATOR_BACKSLASH + Constant.SEPARATOR_DOT)[0];
		String url = "jdbc:dbf:///" + dataBasePath + "?charSet=GBK";
		try {
			Connection con = DriverManager.getConnection(url);
			PreparedStatement stmt = con.prepareStatement("DROP TABLE IF EXISTS" + tableName + ";");
			int a = stmt.executeUpdate();
			logger.info(a + "");
			return a > 0;
		} catch (SQLException e) {
			logger.warn("连接[" + filePath + "]DBF文件异常" + e);
		}
		return false;
	}

}
