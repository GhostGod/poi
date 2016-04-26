package com.xxx.io;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Excel2003文件Reader
 */
public class Excel2003Reader implements HSSFListener, ImportReader {

	private int minColumns = -1;
	private POIFSFileSystem fs;
	private int lastRowNumber;
	private int lastColumnNumber;

	/** Should we output the formula, or the value it has? */
	private boolean outputFormulaValues = true;

	/** For parsing Formulas */
	private SheetRecordCollectingListener workbookBuildingListener;
	//excel2003工作薄
	private HSSFWorkbook stubWorkbook;

	// Records we pick up as we process
	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;

	//表索引
	//private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	@SuppressWarnings("rawtypes")
	private List boundSheetRecords = new ArrayList();

	// For handling formulas with string results
	private int nextRow;
	private int nextColumn;
	private boolean outputNextStringRecord;
	//当前行
	private int curRow = 0;
	//存储行记录的容器
	private Map<String, String> rowData = new HashMap<String, String>();
	private Map<String, String> header = new LinkedHashMap<String, String>();
	private InputStream in;

	private RowReader rowReader;

	public Excel2003Reader(String filename) throws FileNotFoundException {
		this.in = new FileInputStream(filename);
	}

	public Excel2003Reader(InputStream in) {
		this.in = in;
	}

	@Override
	public void setRowReader(RowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**
	 * 遍历excel下所有的sheet
	 * @throws IOException
	 */
	@Override
	public void process() throws IOException {
		this.fs = new POIFSFileSystem(in);
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
		formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, fs);
	}

	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@SuppressWarnings("unchecked")
	@Override
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String value = null;
		switch (record.getSid()) {

		case BoundSheetRecord.sid:
			boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				// 如果有需要，则建立子工作薄
				if (workbookBuildingListener != null && stubWorkbook == null) {
					stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
				}

				//sheetIndex++;
				if (orderedBSRs == null) {
					orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
				}
				//sheetName = orderedBSRs[sheetIndex].getSheetname();
			}
			break;

		case SSTRecord.sid:
			sstRecord = (SSTRecord) record;
			break;

		case BlankRecord.sid:
			BlankRecord brec = (BlankRecord) record;
			curRow = thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			rowData.put(String.valueOf(thisColumn), "");
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), "");
			}
			//rowlist.add(thisColumn, thisStr);
			break;
		case BoolErrRecord.sid: //单元格为布尔类型
			BoolErrRecord berec = (BoolErrRecord) record;
			curRow = thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			value = berec.getBooleanValue() + "";
			rowData.put(String.valueOf(thisColumn), value);
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), "");
			}
			//rowlist.add(thisColumn, thisStr);
			break;

		case FormulaRecord.sid: //单元格为公式类型
			FormulaRecord frec = (FormulaRecord) record;
			curRow = thisRow = frec.getRow();
			thisColumn = frec.getColumn();
			if (outputFormulaValues) {
				if (Double.isNaN(frec.getValue())) {
					// Formula result is a string
					// This is stored in the next record
					outputNextStringRecord = true;
					nextRow = frec.getRow();
					nextColumn = frec.getColumn();
				} else {
					value = formatListener.formatNumberDateCell(frec);
				}
			} else {
				value = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
			}
			rowData.put(String.valueOf(thisColumn), value);
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), value);
			}
			//rowlist.add(thisColumn, thisStr);
			break;
		case StringRecord.sid://单元格中公式的字符串
			if (outputNextStringRecord) {
				// String for formula
				StringRecord srec = (StringRecord) record;
				value = srec.getString();
				curRow = thisRow = nextRow;
				thisColumn = nextColumn;
				outputNextStringRecord = false;
			}
			break;
		case LabelRecord.sid:
			LabelRecord lrec = (LabelRecord) record;
			curRow = thisRow = lrec.getRow();
			thisColumn = lrec.getColumn();
			value = lrec.getValue().trim();
			//value = value.equals("") ? " " : value;
			rowData.put(String.valueOf(thisColumn), value);
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), value);
			}
			//rowlist.add(thisColumn, thisStr);
			break;
		case LabelSSTRecord.sid: //单元格为字符串类型
			LabelSSTRecord lsrec = (LabelSSTRecord) record;
			curRow = thisRow = lsrec.getRow();
			thisColumn = lsrec.getColumn();
			if (sstRecord == null) {
				rowData.put(String.valueOf(thisColumn), "");
				if (curRow == 0) {
					header.put(String.valueOf(thisColumn), "");
				}
				//rowlist.add(thisColumn, " ");
			} else {
				value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
				//value = value.equals("") ? " " : value;
				rowData.put(String.valueOf(thisColumn), value);
				if (curRow == 0) {
					header.put(String.valueOf(thisColumn), value);
				}
				//rowlist.add(thisColumn, thisStr);
			}
			break;
		case NumberRecord.sid: //单元格为数字类型
			NumberRecord numrec = (NumberRecord) record;
			curRow = thisRow = numrec.getRow();
			thisColumn = numrec.getColumn();
			value = formatListener.formatNumberDateCell(numrec).trim();
			//value = value.equals("") ? " " : value;
			// 向容器加入列值
			rowData.put(String.valueOf(thisColumn), value);
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), value);
			}
			//rowlist.add(thisColumn, thisStr);
			break;
		default:
			break;
		}

		// 遇到新行的操作
		if (thisRow != -1 && thisRow != lastRowNumber) {
			lastColumnNumber = -1;
		}

		// 空值的操作
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			curRow = thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			rowData.put(String.valueOf(thisColumn), "");
			if (curRow == 0) {
				header.put(String.valueOf(thisColumn), "");
			}
			//rowlist.add(thisColumn, " ");
		}

		// 更新行和列的值
		if (thisRow > -1)
			lastRowNumber = thisRow;
		if (thisColumn > -1)
			lastColumnNumber = thisColumn;

		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (minColumns > 0) {
				// 列值重新置空
				if (lastColumnNumber == -1) {
					lastColumnNumber = 0;
				}
			}
			lastColumnNumber = -1;
			// 每行结束时， 调用getRows() 方法
			if (rowReader != null) {
				rowReader.dealRowData(curRow, rowData);
			}
			// 清空容器
			rowData.clear();
		}
	}

	@Override
	public int getRowCount() {
		return curRow;
	}

	@Override
	public Map<String, String> getHeader() {
		return header;
	}

	@Override
	public void close() throws IOException {
		in.close();
	}
}
