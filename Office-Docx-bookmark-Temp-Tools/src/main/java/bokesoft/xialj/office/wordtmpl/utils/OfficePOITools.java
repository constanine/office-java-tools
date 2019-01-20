package bokesoft.xialj.office.wordtmpl.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.dom4j.DocumentException;

import com.alibaba.fastjson.JSON;

import bokesoft.xialj.office.wordtmpl.bean.BillUnit;
import bokesoft.xialj.office.wordtmpl.bean.HeaderUnit;
import bokesoft.xialj.office.wordtmpl.bean.TableUnit;

public class OfficePOITools {
	
	public static String readWordToJson(String inputUrl,Map<String,String> headRelatedMap,Map<String,String> dtlRelatedMap) throws DocumentException, IOException, XmlException {
		return readWordToJson(inputUrl, headRelatedMap, dtlRelatedMap, false);
	}
	
	/**
	 * 从指定的docx模板文件中获取单据字段的设计格式的json字符串
	 * @param inputUrl 指定的docx模板文件路径
	 * @return
	 * @throws DocumentException
	 * @throws IOException
	 * @throws XmlException 
	 */
	public static String readWordToJson(String inputUrl,Map<String,String> headRelatedMap,Map<String,String> dtlRelatedMap,boolean editCfg) throws DocumentException, IOException, XmlException {
		BillUnit BillUnit = new BillUnit();
		// 获取word文档解析对象
		XWPFDocument doucument = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
		ArrayList<HeaderUnit> headers = WordParser.transfHeadDatasfromBookmark(doucument,headRelatedMap);
		ArrayList<TableUnit> tables = WordParser.transfDtlDatasfromTable(doucument,dtlRelatedMap,editCfg);
		BillUnit.setHeaders(headers);
		BillUnit.setTables(tables);
		doucument.getPackage().revert();
		return JSON.toJSONString(BillUnit);
	}
	
	/**
	 * 根据Docx模板文件,将传入的数据,按照模板文件的格式,写入到新的docx附件中
	 * @param inputUrl Docx模板文件路径
	 * @param outputUrl ocx附件路径
	 * @param BillUnit 将传入的特定数据
	 * @param showOpTitle 选项是否显示选择标题
	 * @throws Exception
	 */
	public static void writeWordToData(String inputUrl, String outputUrl, BillUnit BillUnit) throws Exception {
		// 获取word文档解析对象
		XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
		List<HeaderUnit> headers = BillUnit.getHeaders();
		List<TableUnit> tables = BillUnit.getTables();
		WordParser.writeHead2Word(document, headers);
		WordParser.writeDtlTable2Word(document, tables);
		FileOutputStream outStream = new FileOutputStream(outputUrl);
		document.write(outStream);
		outStream.close();
		document.getPackage().revert();
	}
}
