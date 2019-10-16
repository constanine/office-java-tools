package bokesoft.xialj.office.wordtmpl.utils;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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

public class WordParser {
	private static final String STARTMARK = "S";
	private static final String ENDMARK = "E";
	private static final String WILDCARD_PARAGRAPH = "</br>";

	public static final String COMP_FIXEDTABLEROW_KEY = "fixTable";
	public static final String COMP_SHOW_KEY = "SHOW";
	public static final String COMP_SELECT_KEY = "OP";

	/**
	 * 获取word中书签
	 * 
	 * @param document
	 * @return
	 * @throws DocumentException
	 * @throws XmlException
	 */
	public static ArrayList<HeaderUnit> transfHeadDatasfromBookmark(XWPFDocument document,
			Map<String, String> relatedMap) throws DocumentException, XmlException {
		// 获取段落文本对象
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		ArrayList<HeaderUnit> result = new ArrayList<HeaderUnit>();
		for (int pgIdx = 0; pgIdx < paragraphs.size(); pgIdx++) {
			XWPFParagraph paragraph = paragraphs.get(pgIdx);
			CTP ctp = paragraph.getCTP();
			List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
			for (int cIdx = 0; cIdx < bookmarks.size();) {
				CTBookmark ctBookmark = bookmarks.get(cIdx);
				String lableName = ctBookmark.getName().toUpperCase();
				if (lableName.indexOf("_") == -1 || lableName.startsWith("_")) {
					cIdx++;
					continue;
				}
				HeaderUnit headDataUnit = new HeaderUnit();
				headDataUnit.setBookMark(lableName);
				if (lableName.startsWith("OP")) {
					int curDataIdx = result.size() - 1;
					headDataUnit = result.get(curDataIdx);
					List<OptionDataUnit> optionList = headDataUnit.getOptionList();
					int optionListLength = 0;
					if (optionList != null) {
						optionListLength = optionList.size();
					} else {
						optionList = new ArrayList<OptionDataUnit>();
					}
					String[] infos = lableName.split("_");
					String startNodeName = lableName;
					if (STARTMARK.equals(infos[4])) {
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
									endNodeName = _checkRule4EndBookMark(startNodeName, endBookmark);
									endParagraph = ebLooParagraph;
									break;
								} else {
									containerParagraphs.add(ebLooParagraph);
									continue;
								}
							}
							descr = _digContentfromParagraphsByBookmarkPos(paragraph, endParagraph, containerParagraphs,
									startNodeName, endNodeName);
						} else {
							CTBookmark endBookmark = bookmarks.get(cIdx);
							endNodeName = _checkRule4EndBookMark(startNodeName, endBookmark);
							descr = _digContentfromSameParagraphByBookmark(paragraph, startNodeName, endNodeName);
						}
						OptionDataUnit optionData = new OptionDataUnit();
						String[] nameAttrStrs = lableName.split("_");
						optionData.setCaption(nameAttrStrs[3]);
						optionData.setKey("OP_" + (optionListLength + 1));
						optionData.setType("OPTION");
						optionData.setDescr(descr);
						optionList.add(optionData);
						headDataUnit.setOptionList(optionList);
					} else {
						cIdx++;
					}
					continue;
				} else if (lableName.indexOf("_OP_") == -1) {
					String[] infos = lableName.split("_");
					headDataUnit.setCaption(infos[0]);
					if (null != relatedMap && null != relatedMap.get(infos[0])) {
						headDataUnit.setKey(relatedMap.get(infos[0]));
					} else {
						headDataUnit.setKey("data_" + infos[1]);
					}
					headDataUnit.setType("TEXT");
					headDataUnit.setRecord("");
				} else {
					String[] infos = lableName.split("_");
					headDataUnit.setCaption(infos[0]);
					if (null != relatedMap && null != relatedMap.get(infos[0])) {
						headDataUnit.setKey(relatedMap.get(infos[0]));
					} else {
						headDataUnit.setKey("data_" + infos[2]);
					}
					headDataUnit.setType("COMOBOBOX");
					headDataUnit.setRecord("");
					List<OptionDataUnit> optionList = new ArrayList<OptionDataUnit>();
					headDataUnit.setOptionList(optionList);

					HeaderUnit showComoboboxUnit = new HeaderUnit();
					showComoboboxUnit.setKey("show_" + infos[2]);
					showComoboboxUnit.setCaption(infos[0] + "内容");
					showComoboboxUnit.setType("SHOW");
					showComoboboxUnit.setRecord("");
					showComoboboxUnit.setBookMark(lableName);
					result.add(showComoboboxUnit);
				}
				result.add(headDataUnit);
				cIdx++;
			}
		}
		return result;
	}

	/**
	 * 获取word中表头
	 * 
	 * @param document
	 * @return
	 */
	public static ArrayList<TableUnit> transfDtlDatasfromTable(XWPFDocument document, 
			Map<String, String> relatedMap,boolean editCfg) {
		// 获取word中所有表格元素
		Iterator<XWPFTable> iterator = document.getTablesIterator();
		XWPFTable table;
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;
		ArrayList<TableUnit> res = new ArrayList<TableUnit>();
		int tabCount = 1;
		while (iterator.hasNext()) {
			table = iterator.next();
			rows = table.getRows();
			TableUnit tableDataUnit = new TableUnit();
			ArrayList<RowUnit> rowList = new ArrayList<RowUnit>();
			tableDataUnit.setKey("dtl" + tabCount);
			if (rows.size() > 0) {
				RowUnit dataRowUnit = new RowUnit();
				dataRowUnit.setRowType(RowUnit.TYPE_DATAROW);
				cells = rows.get(0).getTableCells();
				dataRowUnit.setCollist(_getTableHeadList(tabCount, cells, relatedMap,editCfg));
				rowList.add(dataRowUnit);
				tableDataUnit.setRowlist(rowList);

				for (int rowIdx = 2; rowIdx < rows.size(); rowIdx++) {
					RowUnit fixedRowUnit = new RowUnit();
					fixedRowUnit.setRowType(RowUnit.TYPE_FIXEDROW);
					List<XWPFTableCell> frCells = rows.get(rowIdx).getTableCells();
					fixedRowUnit.setCollist(_getTableFixedList(tabCount, rowIdx, frCells));
					rowList.add(fixedRowUnit);
					tableDataUnit.setRowlist(rowList);
				}
			}
			tabCount++;
			res.add(tableDataUnit);
		}
		return res;
	}

	/**
	 * 
	 * @param document
	 * @param headDataUnitList
	 *            头表信息
	 * @param showOpTitle
	 *            显示选项标题
	 * @throws DocumentException
	 * @throws XmlException
	 */
	public static void writeHead2Word(XWPFDocument document, List<HeaderUnit> headDataUnitList)
			throws DocumentException, XmlException {
		_delContentBetweenBookmarks(document);
		// 获取段落文本对象
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
	public static void writeDtlTable2Word(XWPFDocument document, List<TableUnit> tableDataUnitList) {
		// 获取所有表格
		List<XWPFTable> tables = document.getTables();
		// 数据合法性判断(表格数量)
		if (null != tableDataUnitList) {
			if (tables.size() == tableDataUnitList.size()) {
				for (int i = 0; i < tables.size(); i++) {
					XWPFTable table = tables.get(i);
					List<RowUnit> rowDataUnitList = tableDataUnitList.get(i).getRowlist();
					for (int j = rowDataUnitList.size() - 1; j >= 0; j--) {
						if (!rowDataUnitList.get(j).getRowType().equals(RowUnit.TYPE_DATAROW)) {
							continue;
						}
						List<ColumnUnit> dataRowColumnList = rowDataUnitList.get(j).getCollist();
						// 样板行货获取
						XWPFTableRow tmpRow = table.getRow(1);
						table.addRow(tmpRow, 2);
						// 获取到刚刚插入的行
						XWPFTableRow dataRow = table.getRow(1);
						List<XWPFTableCell> xwpfDataRowCells = dataRow.getTableCells();
						// 判断数据合法性（列）
						if (xwpfDataRowCells.size() != dataRowColumnList.size() - 1) {
							continue;
						}
						for (int k = 0; k < dataRowColumnList.size() - 1; k++) {
							List<XWPFParagraph> paras = xwpfDataRowCells.get(k).getParagraphs();
							for (XWPFParagraph xwpfParagraph : paras) {
								// 设置单元格内容
								_replaceInPara(xwpfParagraph, dataRowColumnList.get(k).getRecord());
							}
						}
						rowDataUnitList.remove(j);
					}

					int lastRow = table.getRows().size() - 1;
					int removeRowIdx = 1;
					for (int j = rowDataUnitList.size() - 1; j >= 0 && lastRow > 1; j--) {
						if (rowDataUnitList.get(j).getRowType().equals(RowUnit.TYPE_DATAROW)) {
							continue;
						}
						// 获取固定行
						XWPFTableRow fixedRow = table.getRow(lastRow);
						List<ColumnUnit> fixedRowColList = rowDataUnitList.get(j).getCollist();
						List<XWPFTableCell> xwpfFixedRowColList = fixedRow.getTableCells();
						// 判断数据合法性（列）
						if (xwpfFixedRowColList.size() != fixedRowColList.size() - 1) {
							continue;
						}
						for (int k = 0; k < fixedRowColList.size() - 1; k++) {
							List<XWPFParagraph> paras = xwpfFixedRowColList.get(k).getParagraphs();
							for (XWPFParagraph xwpfParagraph : paras) {
								// 设置单元格内容
								_replaceInPara(xwpfParagraph,
										fixedRowColList.get(k).getCaption() + fixedRowColList.get(k).getRecord());
							}
						}
						removeRowIdx++;
						lastRow--;
					}
					table.removeRow(table.getRows().size() - removeRowIdx);
				}
			}
		}
	}

	private static void _fillHeadControllerData4TextType(XWPFParagraph paragraph, List<HeaderUnit> headDataUnitList,
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
					if (headDataUnit.getBookMark().equals(nameAttrStr) && ("TEXT".equals(headDataUnit.getType()))) {
						recordList.add(headDataUnit.getRecord());
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

	/**
	 * 
	 * @param document
	 * @param pIdx
	 *            当前段落index
	 * @param headDataUnitList
	 *            头表单元集合
	 * @param nodes
	 *            书签所在的段落
	 * @param showOpTitle
	 *            显示选项标题
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private static void _fillHeadControllerData4ComboboxType(XWPFDocument document, int pIdx,
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
							&& ("COMOBOBOX".equals(headDataUnit.getType()))) {
						for (HeaderUnit showheadDataUnit : headDataUnitList) {
							if ("SHOW".equals(showheadDataUnit.getType())
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
					removeParagraph(document, paragraph);
				}
			} else {
				String[] texts = context.split(WILDCARD_PARAGRAPH);
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

	private static XWPFParagraph truncateParagraphVarXmlWay(XWPFParagraph sourceParagraph, int insertParagraphIdx,
			int start, int end) throws DocumentException, XmlException {
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

	private static XWPFParagraph copyParagraphVarXmlWay(XWPFDocument document, XWPFParagraph sourceParagraph,
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

	private static XWPFParagraph _insertNewParagraphByIndex(XWPFDocument document, int insertParagraphIdx, String text,
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

	private static void _paragraphFormatting(XWPFParagraph paragraph) {
		_paragraphFormatting(paragraph, null);
	}

	private static void _paragraphFormatting(XWPFParagraph paragraph, CTPPr sourceStyle) {
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
	 * 删除下拉书签子项S,E之间的RUN
	 * 
	 * @param document
	 * @throws DocumentException
	 * @throws XmlException
	 */
	private static void _delContentBetweenBookmarks(XWPFDocument document) throws DocumentException, XmlException {
		// 获取段落文本对象
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		List<XWPFParagraph> removeParagraphList = new ArrayList<XWPFParagraph>();
		for (int pgIdx = 0; pgIdx < paragraphs.size(); pgIdx++) {
			XWPFParagraph paragraph = paragraphs.get(pgIdx);
			CTP ctp = paragraph.getCTP();
			List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
			for (int cIdx = 0; cIdx < bookmarks.size();) {
				CTBookmark ctBookmark = bookmarks.get(cIdx);
				String lableName = ctBookmark.getName().toUpperCase();
				if (lableName.indexOf("_") == -1 || lableName.startsWith("_")) {
					cIdx++;
					continue;
				}

				if (lableName.startsWith("OP")) {
					String[] infos = lableName.split("_");
					String startNodeName = lableName;
					if (STARTMARK.equals(infos[4])) {
						String endNodeName = null;
						List<XWPFParagraph> containerParagraphs = new ArrayList<XWPFParagraph>();

						cIdx++;
						if (cIdx >= bookmarks.size()) {
							XWPFParagraph endParagraph = null;
							for (int ebLoopId = (pgIdx + 1); ebLoopId < paragraphs.size(); ebLoopId++) {
								XWPFParagraph ebLooParagraph = paragraphs.get(ebLoopId);
								List<CTBookmark> ebLooPBookMarks = ebLooParagraph.getCTP().getBookmarkStartList();
								if (ebLooPBookMarks.size() > 0) {
									CTBookmark endBookmark = ebLooPBookMarks.get(0);
									endNodeName = _checkRule4EndBookMark(startNodeName, endBookmark);
									endParagraph = ebLooParagraph;
									break;
								} else {
									containerParagraphs.add(ebLooParagraph);
									continue;
								}
							}
							RemoveParagraphInfo removeInfo = _delContentfromParagraphsByBookmarkPos(document, paragraph,
									endParagraph, containerParagraphs, startNodeName, endNodeName);
							if (removeInfo.removeBegin) {
								removeParagraphList.add(paragraph);
							}
							if (removeInfo.removeEnd) {
								removeParagraphList.add(endParagraph);
							}
						} else {
							CTBookmark endBookmark = bookmarks.get(cIdx);
							endNodeName = _checkRule4EndBookMark(startNodeName, endBookmark);
							Boolean removeThis = _delContentfromSameParagraphByBookmark(paragraph, startNodeName,
									endNodeName);
							if (removeThis) {
								removeParagraphList.add(paragraph);
							}
						}
					}
				}
				cIdx++;
			}
		}
		for (XWPFParagraph removeParagraph : removeParagraphList) {
			removeParagraph(document, removeParagraph);
		}
	}

	/**
	 * 删除下拉选项S,E之间的RUN(不在一个段落中)
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
	private static RemoveParagraphInfo _delContentfromParagraphsByBookmarkPos(XWPFDocument document,
			XWPFParagraph startParagraph, XWPFParagraph endParagraph, List<XWPFParagraph> containerParagraphs,
			String startNodeName, String endNodeName) throws DocumentException, XmlException {
		// 有S标记段落的paragraph,定位S标记的位置,S以下的内容加入sbuff中,
		// 定位S标记书签的位置
		RemoveParagraphInfo result = new RemoveParagraphInfo();
		int startPos = _calculatorManualNodePoistion(startNodeName,
				_parseParagraphXml2NodeCollection(startParagraph));

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
			removeParagraph(document, containerParagraph);
		}

		// 有E标记段落的paragraph,定位S标记的位置,E以上的内容加入sbuff中,
		// 定位E标记书签的位置
		int endPos = _calculatorManualNodePoistion(endNodeName,
				_parseParagraphXml2NodeCollection(endParagraph));

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

	private static void removeParagraph(XWPFDocument document, XWPFParagraph containerParagraph) {
		List<IBodyElement> bodyElements = document.getBodyElements();
		for (int bodyIdx = 0; bodyIdx < bodyElements.size(); bodyIdx++) {
			if (bodyElements.get(bodyIdx).equals(containerParagraph)) {
				document.removeBodyElement(bodyIdx);
				break;
			}
		}
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
	private static boolean _delContentfromSameParagraphByBookmark(XWPFParagraph paragraph, String startNodeName,
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
	private static String _digContentfromParagraphsByBookmarkPos(XWPFParagraph paragraph, XWPFParagraph emParagraph,
			List<XWPFParagraph> containerParagraphs, String startNodeName, String endNodeName)
			throws DocumentException, XmlException {
		StringBuffer sbuff = new StringBuffer();
		// S以下的内容加入sbuff中
		// 定位S标记书签的位置
		int startPos = _calculatorManualNodePoistion(startNodeName,
				_parseParagraphXml2NodeCollection(paragraph));
		List<XWPFRun> runs = paragraph.getRuns();
		for (int index = startPos; index < runs.size(); index++) {
			sbuff.append(runs.get(index).text());
		}
		sbuff.append(WILDCARD_PARAGRAPH);
		// 所有包含段落直接写入
		for (XWPFParagraph containerParagraph : containerParagraphs) {
			List<XWPFRun> contRuns = containerParagraph.getRuns();
			for (int index = 0; index < contRuns.size(); index++) {
				sbuff.append(contRuns.get(index).text());
			}
			sbuff.append(WILDCARD_PARAGRAPH);
		}

		// 有E标记段落的paragraph,定位S标记的位置,E以上的内容加入sbuff中
		// 定位E标记书签的位置
		int emEndPos = _calculatorManualNodePoistion(endNodeName,
				_parseParagraphXml2NodeCollection(emParagraph));

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
	private static String _digContentfromSameParagraphByBookmark(XWPFParagraph paragraph, String startNodeName,
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
	private static SameParagraphBookMarkInfo _buildSameParagraphBookMarkInfo(XWPFParagraph paragraph,
			String startNodeName, String endNodeName) throws DocumentException, XmlException {
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
	private static String _checkRule4EndBookMark(String startNodeName, CTBookmark endBookmark) {
		String[] startInfos = startNodeName.split("_");
		String endNodeName;
		endNodeName = endBookmark.getName().toUpperCase();
		String[] endInfos = endNodeName.split("_");
		if (!(ENDMARK.equals(endInfos[4]) && startInfos[0].equals(endInfos[0])) && startInfos[1].equals(endInfos[1])
				&& startInfos[2].equals(endInfos[2]) && startInfos[3].equals(endInfos[3])) {
			throw new RuntimeException("模板设计有误,书签设计中," + startNodeName + ",应该紧跟相应E标签");
		}
		return endNodeName;
	}

	/**
	 * 替换段落里面的变量
	 * 
	 * @param para
	 *            要替换的段落
	 * @param params
	 *            参数
	 */
	private static void _replaceInPara(XWPFParagraph para, String text) {
		List<XWPFRun> runs = para.getRuns();
		while (runs.size() > 0) {
			// 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
			// 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
			para.removeRun(0);
		}
		XWPFRun run = para.insertNewRun(0);
		run.setText(text);
		run.setFontSize(14);
	}

	/**
	 * 获取表格所有列名
	 * 
	 * @param tabCount
	 * @param cells
	 * @return
	 */
	private static List<ColumnUnit> _getTableHeadList(int tabCount, List<XWPFTableCell> cells,
			Map<String, String> relatedMap,boolean editCfg) {
		List<ColumnUnit> res = new ArrayList<ColumnUnit>();
		for (int i = 0; i < cells.size(); i++) {
			ColumnUnit columnDataUnit = new ColumnUnit();
			String caption = cells.get(i).getText().trim();
			columnDataUnit.setCaption(caption);
			int colIdx = i + 1;
			if (null != relatedMap && null != relatedMap.get(caption)) {
				columnDataUnit.setKey(relatedMap.get(caption));
			} else {
				columnDataUnit.setKey("tab_" + tabCount + "_" + colIdx);
			}
			columnDataUnit.setRecord("");
			res.add(columnDataUnit);
		}
		if(editCfg){
			ColumnUnit columnDataUnit = new ColumnUnit();
			columnDataUnit.setCaption("editEnable");
			columnDataUnit.setKey("editEnable");
			columnDataUnit.setRecord("1");
			res.add(columnDataUnit);
		}
		return res;
	}

	/**
	 * 获取表格所有列名
	 * 
	 * @param tabCount
	 * @param cells
	 * @return
	 */
	private static List<ColumnUnit> _getTableFixedList(int tabCount, int rowIdx, List<XWPFTableCell> cells) {
		List<ColumnUnit> res = new ArrayList<ColumnUnit>();
		for (int i = 0; i < cells.size(); i++) {
			ColumnUnit columnDataUnit = new ColumnUnit();
			String caption = cells.get(i).getText().trim();
			columnDataUnit.setCaption(caption);
			int colIdx = i + 1;
			columnDataUnit.setKey(COMP_FIXEDTABLEROW_KEY + "_" + tabCount + "_" + rowIdx + "_" + colIdx);
			columnDataUnit.setRecord("");
			res.add(columnDataUnit);
		}
		ColumnUnit columnDataUnit = new ColumnUnit();
		columnDataUnit.setCaption("editEnable");
		columnDataUnit.setKey("editEnable");
		columnDataUnit.setRecord("1");
		res.add(columnDataUnit);
		return res;
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
	private static List<org.w3c.dom.Node> _parseParagraphXml2NodeCollection(XWPFParagraph paragraph)
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
	private static int _calculatorManualNodePoistion(String nodeName, List<org.w3c.dom.Node> nodes) {
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
	private static Map<String, Integer> _calculatorManualNodePoistions(List<String> nodeNames,
			List<org.w3c.dom.Node> nodes) {
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

class SameParagraphBookMarkInfo {
	int startPos;
	int endPos;
}

class RemoveParagraphInfo {
	boolean removeBegin = false;
	boolean removeEnd = false;
}
