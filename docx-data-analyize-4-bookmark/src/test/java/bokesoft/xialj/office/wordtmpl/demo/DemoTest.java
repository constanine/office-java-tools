package bokesoft.xialj.office.wordtmpl.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.io.IOUtils;
import org.apache.xmlbeans.XmlException;
import org.dom4j.DocumentException;

import com.alibaba.fastjson.JSON;

import bokesoft.xialj.office.wordtmpl.bean.BillUnit;
import bokesoft.xialj.office.wordtmpl.utils.OfficePOITools;

public class DemoTest {
	public static void main(String[] args) throws Exception {
		String sourceUrl="src/test/resources/项目设计合同.docx";
		String targetUrl="src/test/resources/项目设计合同-test.docx";
		String targetUrl2="src/test/resources/项目设计合同-test2.docx";
		String jsonPath="src/test/resources/demoData2.json";
		parserDocx(sourceUrl);
		writeDocxByJson(sourceUrl,targetUrl,jsonPath);
		writeDocxByJson2(sourceUrl,targetUrl2,jsonPath);
	}
	
	public static String parserDocx(String url) throws DocumentException, IOException, XmlException{
		File tmplFile = new File(url);
		String tmplFilePath = tmplFile.getAbsolutePath();
		String jsonStr = OfficePOITools.readWordToJson(tmplFilePath, null, null);
		System.out.println(jsonStr);
		return jsonStr;
	}
	
	public static void writeDocxByJson(String sourceUrl,String targetUrl,String jsonPath) throws Exception{
		File tmplFile = new File(sourceUrl);
		String tmplFilePath = tmplFile.getAbsolutePath();
		File targetFile = new File(targetUrl);
		String targetFilePath = targetFile.getAbsolutePath();
		String jsonStr = IOUtils.toString(new FileInputStream(new File(jsonPath)),"UTF-8");
		BillUnit billUnit = JSON.parseObject(jsonStr, BillUnit.class);
		OfficePOITools.writeWordToData(tmplFilePath, targetFilePath, billUnit);
	}
	
	public static void writeDocxByJson2(String sourceUrl,String targetUrl,String jsonPath) throws Exception{
		File tmplFile = new File(sourceUrl);
		String tmplFilePath = tmplFile.getAbsolutePath();
		File targetFile = new File(targetUrl);
		String targetFilePath = targetFile.getAbsolutePath();
		String jsonStr = IOUtils.toString(new FileInputStream(new File(jsonPath)),"UTF-8");
		BillUnit billUnit = JSON.parseObject(jsonStr, BillUnit.class);
		OfficePOITools.writeWordToData(tmplFilePath, targetFilePath, billUnit);
	}
}
