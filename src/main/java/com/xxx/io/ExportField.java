package com.xxx.io;

/**
 * 导出字段辅助类
 */
public class ExportField {
	/**
	 * 导出列属性名称
	 */
	private String name;
	/**
	 * 显示名称
	 */
	private String showName;
	/**
	 * 类型
	 */
	private DataType type;
	/**
	 * 长度
	 */
	private int length;

	/**
	 * 构造方法（导出excel，csv）
	 * @param name 属性名称
	 * @param showName 显示名称
	 * @param type 类型
	 * @param length 长度
	 */
	public ExportField(String name, String showName, DataType type) {
		this.name = name;
		this.showName = showName;
		this.type = type;
	}

	/**
	 * 构造方法（导出dbf）
	 * @param name 属性名称
	 * @param showName 显示名称
	 * @param type 类型
	 * @param length 长度
	 */
	public ExportField(String name, String showName, DataType type, int length) {
		this.name = name;
		this.showName = showName;
		this.type = type;
		this.length = length;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getShowName() {
		return showName;
	}

	public void setShowName(String showName) {
		this.showName = showName;
	}

	public DataType getType() {
		return type;
	}

	public void setType(DataType type) {
		this.type = type;
	}

	public int getLength() {
		return length;
	}

	public void setLength(int length) {
		this.length = length;
	}

	/**
	 * 导出数据类型
	 * @author liuyang
	 * @date 2014年10月27日 下午4:03:43
	 */
	public enum DataType {
		/**
		 * 整型
		 */
		INT,
		/**
		 * 百分比
		 */
		PERCENT,
		/**
		 * 布尔型
		 */
		BOOLEAN,
		/**
		 * 字符型
		 */
		STRING,
		/**
		 * 日期型
		 */
		DATE,
		/**
		 * Map类型
		 */
		MAP
	}
}
