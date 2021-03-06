package bokesoft.xialj.office.wordtmpl.demo;

import java.io.IOException;

import org.apache.xmlbeans.XmlException;
import org.dom4j.DocumentException;

import bokesoft.xialj.office.wordtmpl.utils.OfficePOITools;

public class OfficePOIToolsTest {
	public static void testCorrectReadWordCfgJson() throws DocumentException, IOException, XmlException {
		String sourceUrl="src/test/resources/学期成绩单.docx";
		String result = OfficePOITools.INSTANCE.readWordToJson(sourceUrl);
		System.out.println(result);
	}
	
	public static void testError1ReadWordCfgJson() throws DocumentException, IOException, XmlException {
		String sourceUrl="src/test/resources/学期成绩单-错误示范-正文书签设置格式不对.docx";
		String result = OfficePOITools.INSTANCE.readWordToJson(sourceUrl);
		System.out.println(result);
	}
	
	public static void main (String[] args) throws DocumentException, IOException, XmlException {
		testCorrectReadWordCfgJson();
		testError1ReadWordCfgJson();
	}
}
