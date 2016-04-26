package com.xxx.contant;

/**
 * 常量
 * @author liuyang
 */
public final class Constant {
	/**
	 * 批量导出行数
	 */
	public static int BATCH_EXPORT_ROW = 200;
	/**
	 * 空字符串
	 */
	public static final String BLANK_STRING = "";
	/**
	 * 院系名称默认值
	 */
	public static final String DEFAULT_DEPARTMENT = "其他院系";
	/**
	 * 起始年默认值
	 */
	public static final int DEFAULT_BEGIN_YEAR = 1999;
	/**
	 * 结束年默认值
	 */
	public static final int DEFAULT_END_YEAR = 2999;
	/**
	 * 逗号分隔符
	 */
	public static final String SEPARATOR_COMMA = ",";
	/**
	 * 点分隔符
	 */
	public static final String SEPARATOR_DOT = ".";
	/**
	 * 冒号分隔符
	 */
	public static final String SEPARATOR_COLON = ":";
	/**
	 * 斜线分隔符
	 */
	public static final String SEPARATOR_SLASH = "/";
	/**
	 * 反斜线分隔符
	 */
	public static final String SEPARATOR_BACKSLASH = "\\";
	/**
	 * session关键字：导出文件路径
	 */
	public static final String KEY_EXPORT_FILE_PATH = "exportFilePath";
	/**
	 * session关键字：上传文件路径
	 */
	public static final String KEY_UPLOAD_FILE_PATH = "importFilePath";
	/**
	 * session关键字：导入总数
	 */
	public static final String KEY_IMPORT_ROW_NUM = "importRowNum";

	/**
	 * yyyyMMdd正则表达式
	 */
	public static final String REGEXP_DATE = "(([0-9]{3}[1-9]|[0-9]{2}[1-9][0-9]{1}|[0-9]{1}[1-9][0-9]{2}|[1-9][0-9]{3})(((0[13578]|1[02])(0[1-9]|[12][0-9]|3[01]))|((0[469]|11)(0[1-9]|[12][0-9]|30))|(02(0[1-9]|[1][0-9]|2[0-8]))))|((([0-9]{2})(0[48]|[2468][048]|[13579][26])|((0[48]|[2468][048]|[3579][26])00))0229)";

}
