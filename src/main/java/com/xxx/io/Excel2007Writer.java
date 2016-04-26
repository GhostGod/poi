package com.xxx.io;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.xxx.contant.Constant;
import com.xxx.io.ExportField.DataType;

/**
 * excel2007Writer
 */
public class Excel2007Writer extends AbstractExportWriter {

	private static Logger logger = LoggerFactory.getLogger(Excel2007Writer.class);

	private SXSSFWorkbook wb;
	private Sheet sh;
	//文件路径，包含后缀名
	private String filePath;

	private int columns;
	private int rows;
	private CellStyle cellStyle;
	private CellStyle percentStyle;

	public Excel2007Writer(String filePath, String sheetName, List<ExportField> fields) {
		this.wb = new SXSSFWorkbook(Constant.BATCH_EXPORT_ROW);
		this.sh = wb.createSheet(sheetName);
		this.fields = fields;
		this.filePath = filePath;
		this.columns = fields.size();
	}

	public Excel2007Writer(String filePath, String sheetName, List<ExportField> fields, boolean hasMessage) {
		this(filePath, sheetName, fields);
		this.hasMessage = hasMessage;
		if (hasMessage) {
			cellStyle = wb.createCellStyle();
			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
		}
	}

	@Override
	public void initHeader() {
		Row row = sh.createRow(rows++);
		boolean hasPercent = false;
		for (int i = 0; i < columns; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(fields.get(i).getShowName());
			if (fields.get(i).getType().equals(DataType.PERCENT)) {
				hasPercent = true;
			}
		}
		if (hasMessage) {
			Cell cell = row.createCell(columns);
			cell.setCellValue(MESSAGE);
			cell.setCellStyle(cellStyle);
			cell.setAsActiveCell();
		}
		if (hasPercent) {
			percentStyle = wb.createCellStyle();
			percentStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
		}
	}

	@Override
	public void appendData(List<?> data) throws Exception {
		for (Object o : data) {
			appendData(o, null);
		}
	}

	/**
	* 追加数据
	* @param dataList 数据list
	 * @throws Exception 
	*/
	public void appendData(Map<?, String> dataList) throws Exception {
		for (Entry<?, String> entry : dataList.entrySet()) {
			appendData(entry.getKey(), entry.getValue());
		}
	}

	@Override
	public void close() throws IOException {
		FileOutputStream out = new FileOutputStream(filePath);
		wb.write(out);
		out.close();
		wb.dispose();
		logger.info("xlsx保存成功！");
	}

	/**
	 * 追加一条数据
	 * @throws Exception 
	 */
	private void appendData(Object data, String message) throws Exception {
		Row row = sh.createRow(rows++);
		Object[] objs = parseData(data, message);
		for (int i = 0, size = columns; i < size; i++) {
			DataType dataType = fields.get(i).getType();
			Cell cell = null;
			switch (dataType) {
			case STRING:
				cell = row.createCell(i, Cell.CELL_TYPE_STRING);
				cell.setCellValue(objs[i] == null ? "" : objs[i].toString());
				break;
			case INT:
				cell = row.createCell(i, Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(objs[i].toString()));
				break;
			case PERCENT:
				cell = row.createCell(i, Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(objs[i].toString()));
				cell.setCellStyle(percentStyle);
				break;
			case DATE:
				cell = row.createCell(i, Cell.CELL_TYPE_STRING);
				cell.setCellValue((Date) objs[i]);
				break;
			case BOOLEAN:
				cell = row.createCell(i, Cell.CELL_TYPE_BOOLEAN);
				cell.setCellValue(Boolean.parseBoolean(objs[i].toString()));
				break;
			default:
				cell = row.createCell(i, Cell.CELL_TYPE_STRING);
				cell.setCellValue(objs[i] == null ? "" : objs[i].toString());
				break;
			}
		}
		if (hasMessage) {
			//设置背景色
			Cell cell = row.createCell(columns);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(message);
		}
	}

	/**
	 * 追加一条消息
	 */
	public void appendLineMessage(String message) {
		Row row = sh.createRow(rows++);
		Cell cell = row.createCell(0);
		cell.setCellValue(message);
	}
}
