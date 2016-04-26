package com.xxx.io;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;

/**
 * CSV文件Writer
 * @author liuyang
 * @date 2014年10月30日 下午3:59:25
 */
public class CSVWriter extends AbstractExportWriter {

	private BufferedWriter bw;
	private int columns;
	private int rows;

	/**
	 * 构造方法（默认GBK编码）
	 * @param path 文件路径
	 * @param fields 字段
	 * @throws IOException 
	 */
	public CSVWriter(String path, List<ExportField> fields) throws IOException {
		this(path, fields, "GBK");
	}

	/**
	 * 构造方法
	 * @param path 文件路径
	 * @param fields 字段
	 * @param encoding 编码
	 * @throws IOException 
	 */
	public CSVWriter(String path, List<ExportField> fields, String encoding) throws IOException {
		File csv = new File(path); // CSV文件 
		byte[] bom = { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF };
		//防止excel打开csv乱码
		FileOutputStream bcpFileWriter = new FileOutputStream(csv);
		bcpFileWriter.write(bom);
		bcpFileWriter.close();
		//追记模式
		this.bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csv), encoding));
		this.fields = fields;
	}

	@Override
	public void initHeader() throws IOException {
		this.columns = this.fields.size();
		String[] header = null;
		if (hasMessage) {
			header = new String[this.columns + 1];
			header[this.columns] = MESSAGE;
		} else {
			header = new String[this.columns];
		}
		for (int i = 0; i < this.columns; i++) {
			header[i] = this.fields.get(i).getShowName();
		}
		doWriteData(header);
	}

	@Override
	public void appendData(List<?> dataList) throws Exception {
		for (Object o : dataList) {
			appendData(o, null);
		}
	}

	@Override
	public void close() throws IOException {
		if (this.bw != null) {
			this.bw.close();
		}
	}

	/**
	 * 追加一行消息
	 * @throws IOException 
	 */
	public void appendLineMessage(String message) throws IOException {
		this.bw.newLine();
		this.bw.write(message);
		this.rows++;
	}

	/**
	 * 追加一条数据
	 * @throws IOException 
	 */
	private void doWriteData(String[] values) throws IOException {
		StringBuilder sb = new StringBuilder();
		for (int i = 0; i < values.length; i++) {
			sb.append(values[i]).append(",");
		}
		String newline = sb.substring(0, sb.lastIndexOf(","));
		//this.bw.newLine();
		this.bw.write(newline);
		this.rows++;
	}

	public int getRows() {
		return rows;
	}

	public int getColumns() {
		return columns;
	}

	/**
	 * 追加一条数据
	 */
	private void appendData(Object data, String message) throws Exception {
		Object[] objs = parseData(data, message);
		StringBuilder sb = new StringBuilder();
		for (Object obj : objs) {
			sb.append(obj == null ? "" : obj.toString()).append("\t,");
		}
		String newline = sb.substring(0, sb.lastIndexOf(","));
		this.bw.newLine();
		this.bw.write(newline);
		rows++;
	}

}
