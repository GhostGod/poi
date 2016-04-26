package com.xxx.io;

import static com.xxx.contant.Constant.SEPARATOR_BACKSLASH;
import static com.xxx.contant.Constant.SEPARATOR_DOT;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import com.xxx.io.ExportField.DataType;

/**
 * 抽象导出Writer
 */
public abstract class AbstractExportWriter implements ExportWriter {

	public static final String MESSAGE = "";
	/*
	 * 导出字段
	 */
	protected List<ExportField> fields;

	/**
	 * 是否导出消息
	 */
	protected boolean hasMessage;

	/**
	 * 通过反射将数据转换为字符数组
	 * @param data 数据
	 * @param message 消息
	 * @return
	 * @throws Exception 
	 */
	@SuppressWarnings("unchecked")
	public Object[] parseData(Object data, String message) throws Exception {
		int size = fields.size();
		Object[] values = null;
		if (hasMessage) {
			values = new Object[size + 1];
			values[size] = message == null ? "" : message;
		} else {
			values = new Object[size];
		}
		if (data instanceof Map) {
			Map<String, Object> map = (Map<String, Object>) data;
			for (int j = 0; j < size; j++) {
				values[j] = map.get(fields.get(j).getName());
			}
			return values;
		}
		for (int j = 0; j < size; j++) {
			String fieldName = fields.get(j).getName();
			DataType type = fields.get(j).getType();
			Object temp = data;
			if (fieldName.contains(SEPARATOR_DOT)) {
				String[] fieldNames = fieldName.split(SEPARATOR_BACKSLASH + SEPARATOR_DOT);
				for (int k = 0; k < fieldNames.length - 1; k++) {
					temp = getValue(fieldNames[k], type, temp);
					if (temp == null) {
						break;
					}
				}
				fieldName = fieldNames[fieldNames.length - 1];
			}
			if (DataType.MAP.equals(type)) {
				values[j] = getValue(fieldName, type, temp).toString();
			} else {
				values[j] = getValue(fieldName, type, temp);
			}
		}
		return values;
	}

	/**
	 * 根据字段名获取字段值
	 * @param fieldName 字段名称
	 * @param obj 对象
	 * @return
	 * @throws Exception
	 */
	private Object getValue(String fieldName, DataType type, Object obj) throws Exception {
		if (obj == null) {
			return null;
		}
		String methodName = fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
		if (DataType.BOOLEAN.equals(type)) {
			methodName = "is" + methodName;
		} else {
			methodName = "get" + methodName;
		}
		Method m = obj.getClass().getMethod(methodName);
		return m.invoke(obj);
	}
}
