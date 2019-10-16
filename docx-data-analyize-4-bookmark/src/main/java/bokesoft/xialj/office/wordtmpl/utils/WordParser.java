package bokesoft.xialj.office.wordtmpl.utils;

import java.math.BigInteger;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFRun.FontCharRange;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.dom4j.DocumentException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.w3c.dom.NodeList;

import bokesoft.xialj.office.wordtmpl.bean.ColumnUnit;
import bokesoft.xialj.office.wordtmpl.bean.HeaderUnit;
import bokesoft.xialj.office.wordtmpl.bean.OptionDataUnit;
import bokesoft.xialj.office.wordtmpl.bean.RowUnit;
import bokesoft.xialj.office.wordtmpl.bean.TableUnit;
import bokesoft.xialj.office.wordtmpl.type.WordParserCard;
import bokesoft.xialj.office.wordtmpl.type.WordTempComponentType;
import bokesoft.xialj.office.wordtmpl.type.WordTempDataType;
import bokesoft.xialj.office.wordtmpl.type.WordTempRowType;
import bokesoft.xialj.office.wordtmpl.type.WordTempTableType;

public class WordParser {

	public static WordParser INSTANCE = new WordParser();
	public Logger logger = LoggerFactory.getLogger(this.getClass());
	private final DateFormat dateFormat = DateFormat.getDateInstance();
	private final DateFormat timeFormat = DateFormat.getDateTimeInstance();

	/**
	 * 获取word中书签
	 * 
	 * @param document
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	public ArrayList<HeaderUnit> transfHeadDatasfromBookmark(XWPFDocument document, List<String> errorMsgCollection)
			throws DocumentException, XmlException {
		logger.info("start analyze word temp paragraph config");
		Set<String> fieldKeySet = new HashSet<String>();
		ArrayList<HeaderUnit> result = new ArrayList<HeaderUnit>();
		// 获取段落文本对象
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		// 逐段查找正文段落中的书签
		for (int pgIdx = 0; pgIdx < paragraphs.size(); pgIdx++) {
			XWPFParagraph paragraph = paragraphs.get(pgIdx);
			CTP ctp = paragraph.getCTP();
			// 每个段落中存在多个书签的可能性
			List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
			for (int cIdx = 0; cIdx < bookmarks.size();) {

				CTBookmark ctBookmark = bookmarks.get(cIdx);
				String bookmarkStr = ctBookmark.getName().toUpperCase();
				// 存在非解析用书签
				if (!_isAnalyizedCard(bookmarkStr)) {
					cIdx++;
					continue;
				}
				boolean needadd = true;
				String[] filedCfgData = bookmarkStr.split("_");
				boolean hasError = false;
				hasError = _checkBookmarkStrucIsError(errorMsgCollection, bookmarkStr, filedCfgData);
				if (hasError) {
					cIdx++;
					continue;
				}

				String fieldType = filedCfgData[0];
				String filedCfgKey = filedCfgData[1];
				String filedCfgCaption = filedCfgData[2];
				if (fieldType.equals(WordParserCard.BOOKMARK_TYPE_OPTION)) {
					String selectContainerKey = filedCfgData[3];
					int curDataIdx = result.size() - 1;
					HeaderUnit headDataUnit = result.get(curDataIdx);
					if (!headDataUnit.getKey().equals(selectContainerKey)) {
						headDataUnit = _findPatchHeadUnit(result, selectContainerKey);
					}
					if (null == headDataUnit) {
						errorMsgCollection
								.add(_buildHeadErrorMessage(bookmarkStr, "selector[" + selectContainerKey + "]无法有效匹配"));
						hasError = true;
					}
					if (hasError) {
						cIdx++;
						continue;
					}

					String optionKey = "OP_" + selectContainerKey + "_" + filedCfgKey;

					hasError = _checkRepeatKeyError(errorMsgCollection, fieldKeySet, bookmarkStr, optionKey);
					if (hasError) {
						cIdx++;
						continue;
					}

					List<OptionDataUnit> optionList = headDataUnit.getOptionList();
					if (null == optionList) {
						optionList = new ArrayList<OptionDataUnit>();
						headDataUnit.setOptionList(optionList);
					}

					String optionType = filedCfgData[4];
					if (WordParserCard.BOOKMARK_START_CARD.equals(optionType)) {
						String endNodeName = null;
						String descr = null;
						List<XWPFParagraph> containerParagraphs = new ArrayList<XWPFParagraph>();
						cIdx++;
						if (cIdx >= bookmarks.size()) {
							XWPFParagraph endParagraph = null;
							for (int ebLoopId = (pgIdx + 1); ebLoopId < paragraphs.size(); ebLoopId++) {
								XWPFParagraph ebLooParagraph = paragraphs.get(ebLoopId);
								List<CTBookmark> ebLooPBookMarks = ebLooParagraph.getCTP().getBookmarkStartList();
								if (ebLooPBookMarks.size() > 0) {
									CTBookmark endBookmark = ebLooPBookMarks.get(0);
									endNodeName = _checkRule4EndBookMark(filedCfgData, endBookmark);
									endParagraph = ebLooParagraph;
									break;
								} else {
									containerParagraphs.add(ebLooParagraph);
									continue;
								}
							}
							descr = _digContentfromMultiParagraphsByBookmark(paragraph, endParagraph,
									containerParagraphs, bookmarkStr, endNodeName);
						} else {
							CTBookmark endBookmark = bookmarks.get(cIdx);
							endNodeName = _checkRule4EndBookMark(filedCfgData, endBookmark);
							descr = _digContentfromSameParagraphByBookmark(paragraph, bookmarkStr, endNodeName);
						}
						OptionDataUnit optionData = new OptionDataUnit();
						optionData.setCaption(filedCfgCaption);
						optionData.setKey(optionKey);
						optionData.setComponentType(WordTempComponentType.OPTION);
						optionData.setDescr(descr);
						optionList.add(optionData);
					}
				} else if (fieldType.equals(WordParserCard.BOOKMARK_TYPE_SELECT)) {
					hasError = _checkRepeatKeyError(errorMsgCollection, fieldKeySet, bookmarkStr, filedCfgKey);
					if (hasError) {
						cIdx++;
						continue;
					}
					HeaderUnit headDataUnit = new HeaderUnit();
					headDataUnit.setBookMark(bookmarkStr);
					headDataUnit.setKey(filedCfgKey);
					headDataUnit.setCaption(filedCfgCaption);
					headDataUnit.setDataType(WordTempDataType.STRING);
					headDataUnit.setComponentType(WordTempComponentType.COMOBOBOX);
					headDataUnit.setRecord("");
					List<OptionDataUnit> optionList = new ArrayList<OptionDataUnit>();
					headDataUnit.setOptionList(optionList);
					result.add(headDataUnit);

					HeaderUnit showComoboboxUnit = new HeaderUnit();
					showComoboboxUnit.setBookMark(bookmarkStr);
					showComoboboxUnit.setKey("SHOW_" + filedCfgKey);
					showComoboboxUnit.setCaption(filedCfgCaption + "的内容");
					showComoboboxUnit.setDataType(WordTempDataType.STRING);
					showComoboboxUnit.setComponentType(WordTempComponentType.SHOW);
					showComoboboxUnit.setRecord("");
					result.add(showComoboboxUnit);
				} else {
					hasError = _checkRepeatKeyError(errorMsgCollection, fieldKeySet, bookmarkStr, filedCfgKey);
					if (hasError) {
						cIdx++;
						continue;
					}
					HeaderUnit headDataUnit = new HeaderUnit();
					headDataUnit.setBookMark(bookmarkStr);
					headDataUnit.setKey(filedCfgKey);
					headDataUnit.setCaption(filedCfgCaption);
					headDataUnit.setDataType(WordTempDataType.getType(fieldType.toUpperCase()));
					headDataUnit.setComponentType(WordTempComponentType.getType(fieldType.toUpperCase()));
					headDataUnit.setRecord("");
					result.add(headDataUnit);
				}
				if (needadd) {
					cIdx++;
				}
			}
		}
		return result;
	}

	private String _buildHeadErrorMessage(String bookmarkStr, String errReason) {
		return "正文类型书签[" + bookmarkStr + "]设置错误," + errReason;
	}

	private String _buildDtlErrorMessage(String bookmarkStr, String errReason) {
		return "表格类型书签[" + bookmarkStr + "]设置错误," + errReason;
	}

	private boolean _checkBookmarkStrucIsError(List<String> errorMsgCollection, String bookmarkStr,
			String[] filedCfgData) {
		if (filedCfgData[0].equals(WordParserCard.BOOKMARK_TYPE_OPTION)) {
			if (filedCfgData.length != 5) {
				errorMsgCollection.add(_buildHeadErrorMessage(bookmarkStr,
						"结构设置不正确,请按${controlltype}_${key}_${caption}_${SELECTOR}_${S/E}设置"));
				return true;
			}
		} else {
			if (filedCfgData.length != 3) {
				errorMsgCollection
						.add(_buildHeadErrorMessage(bookmarkStr, "结构设置不正确,请按${controlltype}_${key}_${caption}设置"));
				return true;
			}
		}
		return false;
	}

	private boolean _checkRepeatKeyError(List<String> errorMsgCollection, Set<String> fieldKeySet, String bookmarkStr,
			String filedCfgKey) {
		if (fieldKeySet.contains(filedCfgKey)) {
			errorMsgCollection.add(_buildHeadErrorMessage(bookmarkStr, "key[" + filedCfgKey + "]重复"));
			return true;
		}
		return false;
	}

	private boolean _isAnalyizedCard(String lableName) {
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_INT)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_LONG)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_TEXT)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_OPTION)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_SELECT)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_DATE)) {
			return true;
		}
		if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_DATETIME)) {
			return true;
		}
		return false;
	}

	private HeaderUnit _findPatchHeadUnit(List<HeaderUnit> headerUnitList, String patchKey) {
		for (HeaderUnit patchHeadDataUnit : headerUnitList) {
			if (patchHeadDataUnit.getKey().equals(patchKey)) {
				return patchHeadDataUnit;
			}
		}
		return null;
	}

	/**
	 * 获取word中表头
	 * 
	 * @param document
	 * @return
	 */
	public ArrayList<TableUnit> transfDtlDatasfromTable(XWPFDocument document, List<String> errorMsgCollection,
			boolean editCfg) {
		logger.info("start analyze word temp tables config");
		Iterator<XWPFTable> iterator = document.getTablesIterator();
		ArrayList<TableUnit> res = new ArrayList<TableUnit>();
		int tabCount = 0;
		while (iterator.hasNext()) {
			boolean hasError = false;
			tabCount++;
			XWPFTable table = iterator.next();
			List<CTBookmark> talbeBookmarks = table.getRow(0).getCell(0).getParagraphs().get(0).getCTP()
					.getBookmarkStartList();
			if (talbeBookmarks.size() > 0) {
				String tableBookmark = _findCellBookmark(talbeBookmarks, WordParserCard.BOOKMARK_TYPE_TABLE);
				String[] tableCfgData = tableBookmark.split("_");
				hasError = _checkTableBookmarkStrucIsError(errorMsgCollection, hasError, tableBookmark, tableCfgData);
				if (hasError) {
					continue;
				}
				String tableTypeStr = tableCfgData[1];
				String tableKey = tableCfgData[2];
				String tableCaption = tableCfgData[3];
				TableUnit tableDataUnit = new TableUnit(tabCount, tableKey, tableCaption);
				tableDataUnit.setRowUnitList(new ArrayList<RowUnit>());
				res.add(tableDataUnit);
				if (WordTempTableType.NORMAL.equals(tableTypeStr)) {
					tableDataUnit.setTableType(WordTempTableType.NORMAL);
					List<XWPFTableRow> rows = table.getRows();
					if (rows.size() > 0) {
						// 普通表格,第一行是数据行
						RowUnit dataRowUnit = new RowUnit();
						tableDataUnit.getRowUnitList().add(dataRowUnit);
						dataRowUnit.setRowType(WordTempRowType.DATA);
						List<XWPFTableCell> dataCells = rows.get(0).getTableCells();
						dataRowUnit.setCollist(_getTableRowColumns(dataCells, errorMsgCollection, editCfg));
						dataRowUnit.setRowIdx(0);
						for (int rowIdx = 2; rowIdx < rows.size(); rowIdx++) {
							XWPFTableRow fixRow = rows.get(rowIdx);
							_constructRowUnit(rowIdx,WordTempRowType.FIXED, tableDataUnit.getRowUnitList(), fixRow,
									errorMsgCollection, editCfg);
						}
					}
				}
				if (WordTempTableType.NOTABLE.equals(tableTypeStr)) {
					String tableType = WordTempTableType.NOTABLE;
					String rowType = WordTempRowType.NOTABLE;
					_analyzeTable4FixOrNotable(tableDataUnit, table, tableType, rowType, errorMsgCollection, editCfg);

				} else if (WordTempTableType.ALLFIXED.equals(tableTypeStr)) {
					String tableType = WordTempTableType.ALLFIXED;
					String rowType = WordTempRowType.FIXED;
					_analyzeTable4FixOrNotable(tableDataUnit, table, tableType, rowType, errorMsgCollection, editCfg);
				}
			}
		}
		return res;
	}

	private boolean _checkTableBookmarkStrucIsError(List<String> errorMsgCollection, boolean hasError, String bookmark,
			String[] bookMarkData) {
		if (bookMarkData.length != 4) {
			errorMsgCollection.add(_buildDtlErrorMessage(bookmark, "结构设置不正确,请按TABLE_${tableType}_${key}_${caption}设置"));
			return true;
		}
		return false;
	}

	/**
	 * 
	 * @param document
	 * @param headDataUnitList 头表信息
	 * @param showOpTitle      显示选项标题
	 * @throws DocumentException
	 * @throws XmlException
	 */
	public void writeHead2Word(XWPFDocument document, List<HeaderUnit> headDataUnitList)
			throws DocumentException, XmlException {
		_delContentBetweenBookmarks(document);
		// 优先将Text的类型填充
		List<XWPFParagraph> paragraphList = document.getParagraphs();
		for (int pIdx = 0; pIdx < paragraphList.size(); pIdx++) {
			XWPFParagraph paragraph = paragraphList.get(pIdx);
			if (paragraph.getCTP().getBookmarkStartList().size() > 0) {
				List<org.w3c.dom.Node> nodes = _parseParagraphXml2NodeCollection(paragraph);
				_fillHeadControllerData4TextType(paragraph, headDataUnitList, nodes);
			}
		}
		for (int pIdx = 0; pIdx < paragraphList.size(); pIdx++) {
			XWPFParagraph paragraph = paragraphList.get(pIdx);
			if (paragraph.getCTP().getBookmarkStartList().size() > 0) {
				List<org.w3c.dom.Node> nodes = _parseParagraphXml2NodeCollection(paragraph);
				_fillHeadControllerData4ComboboxType(document, pIdx, headDataUnitList, nodes);
			}
		}
	}

	/**
	 * 表格文本替换
	 * 
	 * @param document
	 * @param titleLableList
	 */
	public void writeDtlTable2Word(XWPFDocument document, List<TableUnit> tableDataUnitList) {
		// 获取所有表格
		Iterator<XWPFTable> iterator = document.getTablesIterator();
		while (iterator.hasNext()) {
			XWPFTable table = iterator.next();
			List<CTBookmark> talbeBookmarks = table.getRow(0).getCell(0).getParagraphs().get(0).getCTP()
					.getBookmarkStartList();
			if (talbeBookmarks.size() > 0) {
				String tableBookmark = _findCellBookmark(talbeBookmarks, WordParserCard.BOOKMARK_TYPE_TABLE);
				String[] tableCfgData = tableBookmark.split("_");
				String tableTypeStr = tableCfgData[1];
				String tableKey = tableCfgData[2];
				String tableCaption = tableCfgData[3];
				TableUnit fillTableDataUnit = _findReleatedTableDataUnit(tableDataUnitList, tableTypeStr, tableKey,
						tableCaption);
				if (null == fillTableDataUnit) {
					throw new RuntimeException("无法找到表格[" + tableKey + "/" + tableCaption + "]的对应数据");
				}

				if (WordTempTableType.NORMAL.equals(tableTypeStr)) {
					List<RowUnit> dataRowList = _findDataTableRow(fillTableDataUnit.getRowUnitList());
					List<RowUnit> fixedRowList = _findFixedTableRow(fillTableDataUnit.getRowUnitList());
					List<XWPFTableRow> rows = table.getRows();
					for (int rowIdx = 2; rowIdx < rows.size(); rowIdx++) {
						XWPFTableRow fixRow = rows.get(rowIdx);
						_filltFiexdRowUnit(fixedRowList, fixRow,rowIdx);
					}
					int dataRowSize = 0;
					for (RowUnit dataRow : dataRowList) {
						// 样板行货获取
						XWPFTableRow tmpRow = table.getRow(1);
						table.addRow(tmpRow, 2);
						XWPFTableRow fillRow = table.getRow(1);
						List<XWPFTableCell> xwpfDataRowCells = fillRow.getTableCells();
						int colIdx = 0;
						for (XWPFTableCell cell : xwpfDataRowCells) {
							ColumnUnit columnUnit = _findReleatdColumn(dataRow.getCollist(), colIdx);
							if(null != columnUnit) {
								_fillTableColumnData(cell, columnUnit);
							}
							colIdx++;
						}
						dataRowSize++;
					}
					table.removeRow(1+dataRowSize);
					table.removeRow(table.getRows().size()-fixedRowList.size());
				}
				if (WordTempTableType.NOTABLE.equals(tableTypeStr)) {
					String tableType = WordTempTableType.NOTABLE;
					String rowType = WordTempRowType.NOTABLE;
					_fillTable4FixOrNotable(fillTableDataUnit, table, tableType, rowType);

				} else if (WordTempTableType.ALLFIXED.equals(tableTypeStr)) {
					String tableType = WordTempTableType.ALLFIXED;
					String rowType = WordTempRowType.FIXED;
					_fillTable4FixOrNotable(fillTableDataUnit, table, tableType, rowType);
				}
			}
		}
	}

	private void _fillTableColumnData(XWPFTableCell cell, ColumnUnit columnUnit) {
		while (cell.getParagraphs().get(0).getRuns().size() > 0) {
			cell.getParagraphs().get(0).removeRun(0);
		}
		XWPFRun run = cell.getParagraphs().get(0).createRun();
		String showVal = columnUnit.getRecord();
		if (WordTempDataType.DATE.equals(columnUnit.getDataType())) {
			showVal = dateFormat.format(new Date(Long.parseLong(showVal)));
		} else if (WordTempDataType.DATETIME.equals(columnUnit.getDataType())) {
			showVal = timeFormat.format(new Date(Long.parseLong(showVal)));
		}
		run.setText(showVal);
		run.setFontFamily("宋体", FontCharRange.ascii);
		run.setFontSize(14);
		cell.getParagraphs().get(0).addRun(run);
	}

	private ColumnUnit _findReleatdColumn(List<ColumnUnit> collist, int colIdx) {
		for (ColumnUnit column : collist) {
			if (column.getColIdx() == colIdx) {
				return column;
			}
		}
		return null;
	}

	private List<RowUnit> _findDataTableRow(List<RowUnit> rowlist) {
		List<RowUnit> result = new ArrayList<RowUnit>();
		for (int rowIdx=rowlist.size()-1;rowIdx>=0;rowIdx--) {
			RowUnit row = rowlist.get(rowIdx);
			if (WordTempRowType.DATA.equals(row.getRowType())) {
				result.add(row);
			}
		}
		return result;
	}

	private List<RowUnit> _findFixedTableRow(List<RowUnit> rowlist) {
		List<RowUnit> result = new ArrayList<RowUnit>();
		for (RowUnit row : rowlist) {
			if (WordTempRowType.FIXED.equals(row.getRowType())) {
				result.add(row);
			}
		}
		return result;
	}

	private void _fillTable4FixOrNotable(TableUnit tableDataUnit, XWPFTable table, String tableType,
			String rowType) {
		List<XWPFTableRow> rows = table.getRows();
		if (rows.size() > 0) {
			int rowIdx = 0;
			for (XWPFTableRow row : rows) {
				_filltFiexdRowUnit(tableDataUnit.getRowUnitList(), row ,rowIdx);
				rowIdx++;
			}
		}
	}

	private void _filltFiexdRowUnit(List<RowUnit> rowDataList, XWPFTableRow XWPFTableRow,int rowIdx) {
		RowUnit releatedRowData = _findReleatedRowData(rowDataList, rowIdx);
		for (XWPFTableCell cell : XWPFTableRow.getTableCells()) {
			List<CTBookmark> bookmarkList = cell.getParagraphs().get(0).getCTP().getBookmarkStartList();
			if (null != bookmarkList && bookmarkList.size() > 0) {
				String cellBookMark = _findCellBookmark(bookmarkList, WordParserCard.BOOKMARK_TYPE_TABLECELL);
				if (null == cellBookMark) {
					continue;
				}
				String[] cellBookMarkData = cellBookMark.split("_");
				ColumnUnit columnUnit = _findReleatedCellUnit(releatedRowData, cellBookMarkData);
				_fillTableColumnData(cell, columnUnit);
			}
		}
	}

	private RowUnit _findReleatedRowData(List<RowUnit> rowDataList, int rowIdx) {
		for(RowUnit rowData:rowDataList) {
			if(rowData.getRowIdx() == rowIdx) {
				return rowData;
			}
		}
		return null;
	}

	private ColumnUnit _findReleatedCellUnit(RowUnit rowData, String[] cellBookMarkData) {
		for (ColumnUnit cell : rowData.getCollist()) {
			if (cell.getKey().equals(cellBookMarkData[2])) {
				return cell;
			}
		}
		return null;
	}

	private TableUnit _findReleatedTableDataUnit(List<TableUnit> tableDataUnitList, String tableTypeStr,
			String tableKey, String tableCaption) {
		TableUnit result = null;
		for (TableUnit tableDataUnit : tableDataUnitList) {
			if (tableDataUnit.getKey().equals(tableKey) && tableDataUnit.getTableType().equals(tableTypeStr)
					&& tableDataUnit.getCaption().equals(tableCaption)) {
				result = tableDataUnit;
			}
		}
		return result;
	}

	private void _analyzeTable4FixOrNotable(TableUnit tableDataUnit, XWPFTable table, String tableType,
			String rowType, List<String> errorMsgCollection, boolean editCfg) {
		tableDataUnit.setTableType(tableType);
		List<XWPFTableRow> rows = table.getRows();
		if (rows.size() > 0) {
			List<RowUnit> rowList = tableDataUnit.getRowUnitList();
			int rowIdx = 0;
			for (XWPFTableRow row : rows) {
				_constructRowUnit(rowIdx,rowType, rowList, row, errorMsgCollection, editCfg);
				rowIdx++;
			}
		}
	}

	private void _constructRowUnit(int rowIdx,String rowType, List<RowUnit> rowList, XWPFTableRow row,
			List<String> errorMsgCollection, boolean editCfg) {
		RowUnit dataRowUnit = new RowUnit();
		dataRowUnit.setRowType(rowType);
		dataRowUnit.setRowIdx(rowIdx);
		List<XWPFTableCell> cells = row.getTableCells();
		dataRowUnit.setCollist(_getTableRowColumns(cells, errorMsgCollection, editCfg));
		if(dataRowUnit.getCollist().size()>0) {
			rowList.add(dataRowUnit);
		}
	}

	private List<ColumnUnit> _getTableRowColumns(List<XWPFTableCell> cells, List<String> errorMsgCollection,
			boolean editCfg) {
		List<ColumnUnit> result = new ArrayList<ColumnUnit>();
		int colIdx = 0;
		for (XWPFTableCell cell : cells) {
			List<CTBookmark> bookmarkList = cell.getParagraphs().get(0).getCTP().getBookmarkStartList();
			boolean hasError = false;
			if (null != bookmarkList && bookmarkList.size() > 0) {
				String cellBookMark = _findCellBookmark(bookmarkList, WordParserCard.BOOKMARK_TYPE_TABLECELL);
				if (null == cellBookMark) {
					continue;
				}
				String[] cellBookMarkData = cellBookMark.split("_");
				String typeStr = cellBookMarkData[1];
				String key = cellBookMarkData[2];
				String caption = cellBookMarkData[3];
				hasError = _checkTableBookmarkStrucIsError(errorMsgCollection, hasError, cellBookMark,
						cellBookMarkData);
				if (hasError) {
					continue;
				}
				ColumnUnit columnDataUnit = new ColumnUnit();
				columnDataUnit.setCaption(caption);
				columnDataUnit.setKey(key);
				columnDataUnit.setDataType(WordTempDataType.getType(typeStr));
				columnDataUnit.setComponentType(WordTempComponentType.getType(typeStr));
				columnDataUnit.setColIdx(colIdx);
				result.add(columnDataUnit);
			}
			colIdx++;
		}
		if (editCfg) {
			ColumnUnit columnDataUnit = new ColumnUnit();
			columnDataUnit.setCaption("editEnable");
			columnDataUnit.setKey("editEnable");
			columnDataUnit.setRecord("1");
			columnDataUnit.setDataType(WordTempDataType.BOOLEAN);
			columnDataUnit.setComponentType(WordTempComponentType.TEXT);
			result.add(columnDataUnit);
		}
		return result;
	}

	private String _findCellBookmark(List<CTBookmark> bookmarkList, String bookMarkCard) {
		String cellBookMark = null;
		for (CTBookmark bookmark : bookmarkList) {
			if (bookmark.getName().startsWith(bookMarkCard)) {
				cellBookMark = bookmark.getName();
				break;
			}
		}
		return cellBookMark;
	}

	/**
	 * 填写头表文本控件的的内容
	 * 
	 * @param paragraph
	 * @param headDataUnitList
	 * @param nodes
	 */
	private void _fillHeadControllerData4TextType(XWPFParagraph paragraph, List<HeaderUnit> headDataUnitList,
			List<org.w3c.dom.Node> nodes) {
		int rIdx = 0;
		List<Integer> posList = new ArrayList<Integer>();
		List<String> recordList = new ArrayList<String>();
		// 优先将Text的类型填充
		for (int eIdx = 0; eIdx < nodes.size(); eIdx++) {
			org.w3c.dom.Element element = (org.w3c.dom.Element) nodes.get(eIdx);
			if ("w:r".equals(element.getNodeName())) {
				rIdx++;
			}
			if ("w:bookmarkStart".equals(element.getNodeName())) {
				String nameAttrStr = element.getAttribute("w:name").toUpperCase();
				// ColumnDataUnit caption 匹配
				for (int uIdx = 0; uIdx < headDataUnitList.size();) {
					HeaderUnit headDataUnit = headDataUnitList.get(uIdx);
					if (headDataUnit.getBookMark().equals(nameAttrStr) && _isPatchNormalBookMark(headDataUnit)) {
						String recordVal = headDataUnit.getRecord();
						if (WordTempDataType.DATE.equals(headDataUnit.getDataType())) {
							recordVal = dateFormat.format(new Date(Long.parseLong(recordVal)));
						} else if (WordTempDataType.DATETIME.equals(headDataUnit.getDataType())) {
							recordVal = timeFormat.format(new Date(Long.parseLong(recordVal)));
						}
						recordList.add(recordVal);
						headDataUnitList.remove(uIdx);
						if (rIdx == 0) {
							posList.add(rIdx + 1);
						} else {
							posList.add(rIdx);
						}
						break;
					}
					uIdx++;
				}
			}
		}

		for (int i = 0; i < posList.size(); i++) {
			XWPFRun run = paragraph.insertNewRun(posList.get(i) + i);
			if (null == run) {
				run = paragraph.createRun();
			}
			run.setText(recordList.get(i));
			run.setFontFamily("宋体", FontCharRange.ascii);
			run.setFontSize(14);
			_paragraphFormatting(paragraph);
		}
	}

	private boolean _isPatchNormalBookMark(HeaderUnit headDataUnit) {
		if (WordTempComponentType.TEXT.equals(headDataUnit.getComponentType())) {
			return true;
		} else if (WordTempComponentType.LONG.equals(headDataUnit.getComponentType())) {
			return true;
		} else if (WordTempComponentType.NUMBER.equals(headDataUnit.getComponentType())) {
			return true;
		} else if (WordTempComponentType.DATE.equals(headDataUnit.getComponentType())) {
			return true;
		}
		return false;
	}

	/**
	 * 填写头表下拉控件的的内容
	 * 
	 * @param document
	 * @param pIdx             当前段落index
	 * @param headDataUnitList 头表单元集合
	 * @param nodes            书签所在的段落
	 * @param showOpTitle      显示选项标题
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private void _fillHeadControllerData4ComboboxType(XWPFDocument document, int pIdx,
			List<HeaderUnit> headDataUnitList, List<org.w3c.dom.Node> nodes) throws DocumentException, XmlException {
		String context = null;
		// 优先将Text的类型填充
		int eIdx = 0;
		int runIdx = 0;
		boolean needHandle = false;
		loop1: for (; eIdx < nodes.size(); eIdx++) {
			org.w3c.dom.Element element = (org.w3c.dom.Element) nodes.get(eIdx);
			if ("w:r".equals(element.getNodeName())) {
				runIdx++;
			}
			if ("w:bookmarkStart".equals(element.getNodeName())) {
				String nameAttrStr = element.getAttribute("w:name").toUpperCase();
				// ColumnDataUnit caption 匹配
				for (int uIdx = 0; uIdx < headDataUnitList.size();) {
					HeaderUnit headDataUnit = headDataUnitList.get(uIdx);
					if (headDataUnit.getBookMark().equals(nameAttrStr)
							&& (WordTempComponentType.COMOBOBOX.equals(headDataUnit.getComponentType()))) {
						for (HeaderUnit showheadDataUnit : headDataUnitList) {
							if (WordTempComponentType.SHOW.equals(showheadDataUnit.getComponentType())
									&& showheadDataUnit.getBookMark().equals(headDataUnit.getBookMark())) {
								context = showheadDataUnit.getRecord();
								break;
							}
						}
						if (StringUtils.isBlank(context)) {
							List<OptionDataUnit> optionList = headDataUnit.getOptionList();
							for (OptionDataUnit option : optionList) {
								if (headDataUnit.getRecord().equals(option.getKey())) {
									context = option.getDescr();
									break;
								}
							}
						}
						headDataUnitList.remove(uIdx);
						needHandle = true;
						break loop1;
					}
					uIdx++;
				}
			}
		}
		if (needHandle) {
			XWPFParagraph paragraph = document.getParagraphArray(pIdx);
			if (StringUtils.isBlank(context)) {
				while (!"[".equals(paragraph.getRuns().get(runIdx - 1).text())) {
					paragraph.removeRun(runIdx - 1);
				}
				paragraph.removeRun(runIdx - 1);
				while (!"]".equals(paragraph.getRuns().get(runIdx - 1).text())) {
					paragraph.removeRun(runIdx - 1);
				}
				paragraph.removeRun(runIdx - 1);
				if (paragraph.getRuns().size() == 0) {
					_removeParagraph(document, paragraph);
				}
			} else {
				context = context.replace("\r\n",WordParserCard.WILDCARD_PARAGRAPH);
				String[] texts = context.split(WordParserCard.WILDCARD_PARAGRAPH);
				if (texts.length > 1) {
					CTPPr sourceStyle = (CTPPr) paragraph.getCTP().getPPr().copy();
					int insertParagraphIdx = pIdx + 1;
					for (String text : texts) {
						_insertNewParagraphByIndex(document, insertParagraphIdx, text, sourceStyle);
						insertParagraphIdx++;
					}
					copyParagraphVarXmlWay(document, paragraph, insertParagraphIdx, eIdx + 3, -1, sourceStyle);
					truncateParagraphVarXmlWay(paragraph, pIdx, 0, eIdx + 1);
				} else {
					XWPFRun run = paragraph.insertNewRun(runIdx);
					run.setText(context);
					run.setFontFamily("宋体", FontCharRange.ascii);
					run.setFontSize(12);
				}
			}
		}
	}

	/**
	 * 通过xml方式,截取word段落
	 * 
	 * @param sourceParagraph    要删除的段落
	 * @param insertParagraphIdx
	 * @param start
	 * @param end
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private XWPFParagraph truncateParagraphVarXmlWay(XWPFParagraph sourceParagraph, int insertParagraphIdx, int start,
			int end) throws DocumentException, XmlException {
		CTP sctp = sourceParagraph.getCTP();
		String spgXml = sctp.xmlText();
		XmlObject sxml = XmlObject.Factory.parse(spgXml);
		org.w3c.dom.Node rootNode = sxml.getDomNode();
		NodeList children = rootNode.getChildNodes();
		for (int nIdx = children.getLength() - 1; nIdx > end;) {
			org.w3c.dom.Node child = children.item(nIdx);
			if (null != child.getNodeName()) {
				rootNode.removeChild(child);
			}
			nIdx = children.getLength() - 1;
		}
		for (int nIdx = 0; nIdx < start;) {
			org.w3c.dom.Node child = children.item(nIdx);
			if (null != child.getNodeName()) {
				rootNode.removeChild(child);
				nIdx++;
			}
		}
		sctp.set(sxml);
		return sourceParagraph;
	}

	/**
	 * 通过xml方式,复制word段落
	 * 
	 * @param document
	 * @param sourceParagraph    源段落
	 * @param insertParagraphIdx 插入的段落位置
	 * @param start
	 * @param end
	 * @param sourceStyle
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private XWPFParagraph copyParagraphVarXmlWay(XWPFDocument document, XWPFParagraph sourceParagraph,
			int insertParagraphIdx, int start, int end, CTPPr sourceStyle) throws DocumentException, XmlException {
		XmlCursor cursor = document.getParagraphArray(insertParagraphIdx).getCTP().newCursor();
		XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
		CTP tctp = newParagraph.getCTP();

		CTP sctp = sourceParagraph.getCTP();
		String spgXml = sctp.xmlText();
		XmlObject sxml = XmlObject.Factory.parse(spgXml);
		org.w3c.dom.Node sRootNode = sxml.getDomNode();

		NodeList schildren = sRootNode.getChildNodes();
		int effectElementIdx = 0;
		for (int nIdx = 0; nIdx < schildren.getLength();) {
			org.w3c.dom.Node child = schildren.item(nIdx);
			if (null != child.getNodeName()) {
				effectElementIdx++;
				if (!(effectElementIdx >= start && (effectElementIdx <= end || end < 0))) {
					sRootNode.removeChild(child);
				} else {
					nIdx++;
				}
			}
		}
		tctp.set(sxml);
		_paragraphFormatting(newParagraph, sourceStyle);
		return newParagraph;
	}

	/**
	 * 插入段落到指定段落位置
	 * 
	 * @param document
	 * @param insertParagraphIdx
	 * @param text               段落内容
	 * @param sourceStyle        段落显示格式
	 * @return
	 */
	private XWPFParagraph _insertNewParagraphByIndex(XWPFDocument document, int insertParagraphIdx, String text,
			CTPPr sourceStyle) {
		XmlCursor cursor = document.getParagraphArray(insertParagraphIdx).getCTP().newCursor();
		XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
		_paragraphFormatting(newParagraph, sourceStyle);
		if (StringUtils.isNotBlank(text)) {
			XWPFRun run = newParagraph.createRun();
			run.setText(text);
			run.setFontFamily("宋体", FontCharRange.ascii);
			run.setFontSize(12);
		}
		return newParagraph;
	}

	private void _paragraphFormatting(XWPFParagraph paragraph) {
		_paragraphFormatting(paragraph, null);
	}

	/**
	 * 段落显示格式调整
	 * 
	 * @param paragraph
	 * @param sourceStyle
	 */
	private void _paragraphFormatting(XWPFParagraph paragraph, CTPPr sourceStyle) {
		paragraph.setFirstLineIndent(0);
		if (null != sourceStyle) {
			paragraph.getCTP().setPPr((CTPPr) sourceStyle);
		} else {
			CTPPr pPPr = paragraph.getCTP().getPPr();
			CTInd ctInd = pPPr.getInd();
			ctInd.setHanging(new BigInteger("0"));
			ctInd.setHangingChars(new BigInteger("0"));
			ctInd.setLeft(new BigInteger("0"));
			ctInd.setLeftChars(new BigInteger("0"));
			ctInd.setFirstLine(new BigInteger("0"));
			ctInd.setFirstLineChars(new BigInteger("0"));
		}

		paragraph.setIndentationFirstLine(560);
	}

	/**
	 * 删除当前段落
	 * 
	 * @param document
	 * @param containerParagraph
	 */
	private void _removeParagraph(XWPFDocument document, XWPFParagraph containerParagraph) {
		List<IBodyElement> bodyElements = document.getBodyElements();
		for (int bodyIdx = 0; bodyIdx < bodyElements.size(); bodyIdx++) {
			if (bodyElements.get(bodyIdx).equals(containerParagraph)) {
				document.removeBodyElement(bodyIdx);
				break;
			}
		}
	}

	/**
	 * 删除下拉书签子项S,E之间的RUN *
	 * 
	 * @param document
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private void _delContentBetweenBookmarks(XWPFDocument document) throws DocumentException, XmlException {
		// 获取段落文本对象
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		List<XWPFParagraph> removeParagraphList = new ArrayList<XWPFParagraph>();
		for (int pgIdx = 0; pgIdx < paragraphs.size(); pgIdx++) {
			XWPFParagraph paragraph = paragraphs.get(pgIdx);
			CTP ctp = paragraph.getCTP();
			List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
			for (int bkIdx = 0; bkIdx < bookmarks.size();) {
				CTBookmark ctBookmark = bookmarks.get(bkIdx);
				String lableName = ctBookmark.getName().toUpperCase();
				if (_isAnalyizedCard(lableName)) {
					bkIdx++;
					continue;
				}

				if (lableName.startsWith(WordParserCard.BOOKMARK_TYPE_OPTION)) {
					String startNodeName = lableName;
					String[] startBookmarkData = lableName.split("_");
					if (WordParserCard.BOOKMARK_START_CARD.equals(startBookmarkData[4])) {
						String endNodeName = null;
						List<XWPFParagraph> containerParagraphs = new ArrayList<XWPFParagraph>();

						bkIdx++;
						if (bkIdx >= bookmarks.size()) {
							XWPFParagraph endParagraph = null;
							for (int ebLoopId = (pgIdx + 1); ebLoopId < paragraphs.size(); ebLoopId++) {
								XWPFParagraph ebLooParagraph = paragraphs.get(ebLoopId);
								List<CTBookmark> ebLooPBookMarks = ebLooParagraph.getCTP().getBookmarkStartList();
								if (ebLooPBookMarks.size() > 0) {
									CTBookmark endBookmark = ebLooPBookMarks.get(0);
									endNodeName = _checkRule4EndBookMark(startBookmarkData, endBookmark);
									endParagraph = ebLooParagraph;
									break;
								} else {
									containerParagraphs.add(ebLooParagraph);
									continue;
								}
							}
							RemoveParagraphInfo removeInfo = _delContentfromMultiParagraphsByBookmark(document,
									paragraph, endParagraph, containerParagraphs, startNodeName, endNodeName);
							if (removeInfo.removeBegin) {
								removeParagraphList.add(paragraph);
							}
							if (removeInfo.removeEnd) {
								if (!removeParagraphList.contains(endParagraph)) {
									removeParagraphList.add(endParagraph);
								}
							}
						} else {
							CTBookmark endBookmark = bookmarks.get(bkIdx);
							endNodeName = _checkRule4EndBookMark(startBookmarkData, endBookmark);
							Boolean removeThis = _delContentfromSameParagraphByBookmark(paragraph, startNodeName,
									endNodeName);
							if (removeThis) {
								removeParagraphList.add(paragraph);
							}
						}
					}
				}
				bkIdx++;
			}
		}
		for (XWPFParagraph removeParagraph : removeParagraphList) {
			_removeParagraph(document, removeParagraph);
		}
	}

	/**
	 * 删除下拉选项S,E之间的RUN(多个段落,不在同一个段落中)
	 * 
	 * @param startParagraph
	 * @param endParagraph
	 * @param containerParagraphs
	 * @param startNodeName
	 * @param endNodeName
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private RemoveParagraphInfo _delContentfromMultiParagraphsByBookmark(XWPFDocument document,
			XWPFParagraph startParagraph, XWPFParagraph endParagraph, List<XWPFParagraph> containerParagraphs,
			String startNodeName, String endNodeName) throws DocumentException, XmlException {
		// 有S标记段落的paragraph,定位S标记的位置,S以下的内容加入sbuff中,
		// 定位S标记书签的位置
		RemoveParagraphInfo result = new RemoveParagraphInfo();
		int startPos = _calculatorManualNodePoistion(startNodeName, _parseParagraphXml2NodeCollection(startParagraph));

		// S以下的内容加入sbuff中
		List<XWPFRun> runs = startParagraph.getRuns();
		while (runs.size() > startPos) {
			startParagraph.removeRun(startPos);
		}

		// 如果S所在的paragraph没有runs了,则直接删除
		if (startParagraph.getRuns().size() <= 0) {
			result.removeBegin = true;
		}

		// 所有包含段落直接写入
		for (XWPFParagraph containerParagraph : containerParagraphs) {
			_removeParagraph(document, containerParagraph);
		}

		// 有E标记段落的paragraph,定位S标记的位置,E以上的内容加入sbuff中,
		// 定位E标记书签的位置
		int endPos = _calculatorManualNodePoistion(endNodeName, _parseParagraphXml2NodeCollection(endParagraph));

		// E以上的内容加入sbuff中
		while (endPos > 0) {
			endParagraph.removeRun(0);
			endPos--;
		}

		// 如果E所在的paragraph没有runs了,则直接删除
		if (endParagraph.getRuns().size() <= 0) {
			result.removeEnd = true;
		}
		return result;
	}

	/**
	 * 删除下拉选项S,E之间的RUN(在一个段落中)
	 * 
	 * @param paragraph
	 * @param startNodeName
	 * @param endNodeName
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private boolean _delContentfromSameParagraphByBookmark(XWPFParagraph paragraph, String startNodeName,
			String endNodeName) throws DocumentException, XmlException {
		boolean needRemoveWholeParagraph = false;
		SameParagraphBookMarkInfo result = _buildSameParagraphBookMarkInfo(paragraph, startNodeName, endNodeName);

		// S以下,E以上的内容加入sbuff中输出
		for (int index = result.startPos; index < result.endPos; index++) {
			paragraph.removeRun(result.startPos);
		}
		if (paragraph.getRuns().size() <= 0) {
			needRemoveWholeParagraph = true;
		}
		return needRemoveWholeParagraph;
	}

	/**
	 * 挖掘下拉选项数据内容,S和E标记不在一个段落中,而且能包含多个无书签的段落
	 * 
	 * @param paragraph
	 * @param emParagraph
	 * @param containerParagraphs
	 * @param startNodeName
	 * @param endNodeName
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private String _digContentfromMultiParagraphsByBookmark(XWPFParagraph paragraph, XWPFParagraph emParagraph,
			List<XWPFParagraph> containerParagraphs, String startNodeName, String endNodeName)
			throws DocumentException, XmlException {
		StringBuffer sbuff = new StringBuffer();
		// S以下的内容加入sbuff中
		// 定位S标记书签的位置
		int startPos = _calculatorManualNodePoistion(startNodeName, _parseParagraphXml2NodeCollection(paragraph));
		List<XWPFRun> runs = paragraph.getRuns();
		for (int index = startPos; index < runs.size(); index++) {
			sbuff.append(runs.get(index).text());
		}
		sbuff.append(WordParserCard.WILDCARD_PARAGRAPH);
		// 所有包含段落直接写入
		for (XWPFParagraph containerParagraph : containerParagraphs) {
			List<XWPFRun> contRuns = containerParagraph.getRuns();
			for (int index = 0; index < contRuns.size(); index++) {
				sbuff.append(contRuns.get(index).text());
			}
			sbuff.append(WordParserCard.WILDCARD_PARAGRAPH);
		}
		// 有E标记段落的paragraph,定位S标记的位置,E以上的内容加入sbuff中
		// 定位E标记书签的位置
		int emEndPos = _calculatorManualNodePoistion(endNodeName, _parseParagraphXml2NodeCollection(emParagraph));
		// E以上的内容加入sbuff中
		List<XWPFRun> emRuns = emParagraph.getRuns();
		for (int index = 0; index < emEndPos; index++) {
			sbuff.append(emRuns.get(index).text());
		}
		return sbuff.toString();
	}

	/**
	 * 挖掘下拉选项数据内容,S和E标记在一个段落中
	 * 
	 * @param paragraph
	 * @param startNodeName
	 * @param endNodeName
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private String _digContentfromSameParagraphByBookmark(XWPFParagraph paragraph, String startNodeName,
			String endNodeName) throws DocumentException, XmlException {
		StringBuffer sbuff = new StringBuffer();
		SameParagraphBookMarkInfo result = _buildSameParagraphBookMarkInfo(paragraph, startNodeName, endNodeName);

		List<XWPFRun> runs = paragraph.getRuns();
		// S以下,E以上的内容加入sbuff中输出
		for (int index = result.startPos; index < result.endPos; index++) {
			sbuff.append(runs.get(index).text());
		}
		return sbuff.toString();
	}

	/**
	 * 计算S,E的runs对应下标集合
	 * 
	 * @param paragraph
	 * @param startNodeName
	 * @param endNodeName
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private SameParagraphBookMarkInfo _buildSameParagraphBookMarkInfo(XWPFParagraph paragraph, String startNodeName,
			String endNodeName) throws DocumentException, XmlException {
		SameParagraphBookMarkInfo result = new SameParagraphBookMarkInfo();
		List<String> nodeNames = new ArrayList<String>();
		nodeNames.add(startNodeName);
		nodeNames.add(endNodeName);
		Map<String, Integer> poistions = _calculatorManualNodePoistions(nodeNames,
				_parseParagraphXml2NodeCollection(paragraph));
		result.startPos = poistions.get(startNodeName);
		result.endPos = poistions.get(endNodeName);
		return result;
	}

	/**
	 * 检查OP下拉数据标记是否规范
	 * 
	 * @param startNodeName
	 * @param endBookmark
	 * @return
	 */
	private String _checkRule4EndBookMark(String[] startInfos, CTBookmark endBookmark) {
		String endNodeName;
		endNodeName = endBookmark.getName().toUpperCase();
		String[] endInfos = endNodeName.split("_");
		if (!(WordParserCard.BOOKMARK_END_CARD.equals(endInfos[4]) && startInfos[0].equals(endInfos[0]))
				&& startInfos[1].equals(endInfos[1]) && startInfos[2].equals(endInfos[2])
				&& startInfos[3].equals(endInfos[3])) {
			throw new RuntimeException("模板设计有误,书签设计中,\" + startNodeName + \",应该紧跟相应E标签");
		}
		return endNodeName;
	}

	/**
	 * 替换段落里面的变量
	 * 
	 * @param para   要替换的段落
	 * @param params 参数
	 */
	protected void _replaceInPara(XWPFParagraph para, String text) {
		List<XWPFRun> runs = para.getRuns();
		while (runs.size() > 0) {
			// 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
			// 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
			para.removeRun(0);
		}
		XWPFRun run = para.insertNewRun(0);
		run.setText(text);
		run.setFontSize(10);
	}

	/**
	 * 将XWPFParagraph转成xml,再有xml转成dom对象,解析后，找到该段落下所有最外层子节点的集合,以便判断出书签的位置
	 * 
	 * @param paragraph
	 * @param elements
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private List<org.w3c.dom.Node> _parseParagraphXml2NodeCollection(XWPFParagraph paragraph)
			throws DocumentException, XmlException {
		List<org.w3c.dom.Node> nodes = new ArrayList<org.w3c.dom.Node>();
		CTP ctp = paragraph.getCTP();
		String pgXml = ctp.xmlText();
		// 将段落改为xml的dom解析
		XmlObject xmlObj = XmlObject.Factory.parse(pgXml);
		org.w3c.dom.Node rootNode = xmlObj.getDomNode();
		NodeList children = rootNode.getChildNodes();
		for (int nIdx = 0; nIdx < children.getLength(); nIdx++) {
			org.w3c.dom.Node child = children.item(nIdx);
			if (null != child.getNodeName()) {
				nodes.add(child);
			}
		}
		return nodes;
	}

	/**
	 * 获取指定nodename的Node在 集合中的下标值
	 * 
	 * @param nodeName
	 * @param elements
	 * @return
	 */
	private int _calculatorManualNodePoistion(String nodeName, List<org.w3c.dom.Node> nodes) {
		List<String> nodeNames = new ArrayList<String>();
		nodeNames.add(nodeName);
		Integer result = _calculatorManualNodePoistions(nodeNames, nodes).get(nodeName);
		return null == result ? -1 : result;
	}

	/**
	 * 获取指定nodename的Node在 集合中的下标值
	 * 
	 * @param nodeName
	 * @param elements
	 * @return
	 */
	private Map<String, Integer> _calculatorManualNodePoistions(List<String> nodeNames, List<org.w3c.dom.Node> nodes) {
		Map<String, Integer> result = new HashMap<String, Integer>();
		int rIdx = 0;
		for (int eIdx = 0; eIdx < nodes.size(); eIdx++) {
			org.w3c.dom.Node node = nodes.get(eIdx);
			if ("w:r".equals(node.getNodeName())) {
				rIdx++;
			}
			// bookmarkStart不计数,r段落计数,保持与runs集合下标一致
			if ("w:bookmarkStart".equals(node.getNodeName())) {
				if (org.w3c.dom.Node.ELEMENT_NODE == node.getNodeType()) {
					org.w3c.dom.Element element = (org.w3c.dom.Element) node;
					String nameAttr = element.getAttribute("w:name").toUpperCase();
					if (nodeNames.contains(nameAttr)) {
						result.put(nameAttr, rIdx);
						// 一旦匹配到,就从nodeNames集合中移除,如果nodeNames集合为空,则可以直接结束循环
						nodeNames.remove(nameAttr);
						if (nodeNames.size() == 0) {
							break;
						}
					}
				}
			}
		}
		return result;
	}
}

/**
 * 相似段落bean
 */
class SameParagraphBookMarkInfo {
	int startPos;
	int endPos;
}

/**
 * 删除段落缓存数据bean
 */
class RemoveParagraphInfo {
	boolean removeBegin = false;
	boolean removeEnd = false;
}
