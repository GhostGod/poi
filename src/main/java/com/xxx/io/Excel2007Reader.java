package com.xxx.io;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * Excel2007文件Reader
 */
public class Excel2007Reader extends DefaultHandler implements ImportReader {
	//共享字符串表
	private SharedStringsTable sst;
	//上一次的内容
	private String lastContents;
	private boolean nextIsString;

	//private int sheetIndex = -1;
	private Map<String, String> rowData = new HashMap<String, String>();
	private Map<String, String> header = new LinkedHashMap<String, String>();
	//当前行
	private int curRow = 0;
	//当前列
	private int curCol = 1;

	private String lastCol;
	private InputStream in;
	//日期标志
	//private boolean dateFlag;
	//数字标志
	//private boolean numberFlag;

	private boolean isTElement;

	private boolean isEmpty;

	private RowReader rowReader;

	public Excel2007Reader(String filename) throws FileNotFoundException {
		this.in = new FileInputStream(filename);
	}

	public Excel2007Reader(InputStream in) {
		this.in = in;
	}

	@Override
	public void setRowReader(RowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
	 * @param filename
	 * @param sheetId
	 * @throws Exception
	 */
	public void processOneSheet() throws Exception {
		OPCPackage pkg = OPCPackage.open(in);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		Iterator<InputStream> sheets = r.getSheetsData();
		if (sheets.hasNext()) {
			curRow = 0;
			//sheetIndex++;
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	@Override
	public void process() throws Exception {
		processOneSheet();
	}

	/**
	 * 遍历工作簿中所有的电子表格
	 * @param filename
	 * @throws Exception
	 */
	public void processAll() throws Exception {
		OPCPackage pkg = OPCPackage.open(in);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			curRow = 0;
			//sheetIndex++;
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}

	/**
	 * 计算中断列数
	 * @param s1 
	 * @param s2
	 * @return 中断了多少列
	 */
	private int interruptColumn(String last, String current) {
		String[] lasts = parseNumberSystem26(last);
		String[] currents = parseNumberSystem26(current);
		if (lasts[1].equals(currents[1])) {
			return fromNumberSystem26(currents[0]) - fromNumberSystem26(lasts[0]) - 1;
		}
		return 0;
	}

	/**
	 * 解析指定的26进制表示。映射关系：[A-Z] ->[1-26]
	 * @param str
	 * @return
	 */
	private String[] parseNumberSystem26(String str) {
		char[] s = str.toCharArray();
		int i = s.length - 1;
		for (; i >= 0; i--) {
			if (s[i] > '9') {
				break;
			}
		}
		return new String[] { str.substring(0, i + 1), str.substring(i + 1) };
	}

	/**
	 * 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]
	 * @param str
	 * @return
	 */
	private int fromNumberSystem26(String str) {
		int n = 0;
		char[] s = str.toCharArray();
		for (int i = s.length - 1, j = 1; i >= 0; i--, j *= 26) {
			char c = s[i];
			if (c < 'A' || c > 'Z')
				return 0;
			n += (c - 64) * j;
		}
		return n;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		//当rowReader为空curRow=0的时候表示读取标题和行数，提高解析标题速度
		if (rowReader != null || curRow == 0) {
			// c => 单元格
			if ("c".equals(name)) {
				// 如果下一个元素是 SST 的索引，则将nextIsString标记为true
				String cellType = attributes.getValue("t");
				if ("s".equals(cellType)) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
				isEmpty = cellType == null;
				String cellColumn = attributes.getValue("r");
				if (curCol > 1) {
					int column = interruptColumn(lastCol, cellColumn);
					for (int i = 1; i <= column; i++) {
						rowData.put(String.valueOf(curCol), "");
						if (curRow == 0) {
							header.put(String.valueOf(curCol), "");
						}
						curCol++;
					}
				} else if (curCol == 1 && !cellColumn.startsWith("A")) {
					rowData.put(String.valueOf(curCol), "");
					if (curRow == 0) {
						header.put(String.valueOf(curCol), "");
					}
					curCol++;
				}
				lastCol = cellColumn;
			}
			if ("v".equals(name)) {
				isEmpty = false;
			}
			//当元素为t时
			if ("t".equals(name)) {
				isTElement = true;
			} else {
				isTElement = false;
			}

			// 置空
			lastContents = "";
		}
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		//当rowReader为空curRow=0的时候表示读取标题和行数，提高解析标题速度
		if (rowReader != null || curRow == 0) {
			if (isEmpty) {
				rowData.put(String.valueOf(curCol), "");
				if (curRow == 0) {
					header.put(String.valueOf(curCol), "");
				}
				curCol++;
			}

			// 根据SST的索引值的到单元格的真正要存储的字符串
			// 这时characters()方法可能会被调用多次
			if (nextIsString) {
				try {
					int idx = Integer.parseInt(lastContents);
					lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				} catch (Exception e) {
					//无需处理，否则程序处理excel会出错
				}
			}
			//t元素也包含字符串
			if (isTElement) {
				String value = lastContents.trim();
				rowData.put(String.valueOf(curCol), value);
				if (curRow == 0) {
					header.put(String.valueOf(curCol), value);
				}
				//rowlist.add(curCol, value);
				curCol++;
				isTElement = false;
				// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
				// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			} else if ("v".equals(name)) {
				String value = lastContents.trim();
				rowData.put(String.valueOf(curCol), value);
				if (curRow == 0) {
					header.put(String.valueOf(curCol), value);
				}
				//rowlist.add(curCol, value);
				curCol++;
			} else {
				//如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
				if (name.equals("row")) {
					if (rowReader != null) {
						rowReader.dealRowData(curRow, rowData);
					}
					rowData.clear();
					curRow++;
					curCol = 1;
					lastCol = "";
				}
			}
		} else {
			if (name.equals("row")) {
				curRow++;
			}
		}
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		//当rowReader为空curRow=0的时候表示读取标题和行数，提高解析标题速度
		if (rowReader != null || curRow == 0) {//得到单元格内容的值
			lastContents += new String(ch, start, length);
		}
	}

	@Override
	public int getRowCount() {
		return curRow - 1;
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
