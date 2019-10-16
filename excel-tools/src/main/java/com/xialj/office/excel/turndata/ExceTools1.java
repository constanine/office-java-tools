package com.xialj.office.excel.turndata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceTools1 {

	public static void main(String[] args) throws Throwable {
		//transfExcelDataForm(args[0],Boolean.valueOf(args[1]));
		transfExcelDataForm("需转换表格-王露贤.xlsx",false);
	}

	private static void transfExcelDataForm(String filePath,boolean hasTitle) throws Throwable{
		File sourceFile = new File(filePath);
		String targetParentPath = sourceFile.getAbsolutePath();
		targetParentPath = targetParentPath.replaceAll("\\\\", "/");
		targetParentPath = targetParentPath.substring(0, targetParentPath.lastIndexOf("/"));
		String targetFileName = sourceFile.getName().substring(0,sourceFile.getName().lastIndexOf(".xlsx"));
		targetFileName += "-转型后";
		File targetFile = new File(targetParentPath+"/"+targetFileName+".xlsx");		
		Map<String,Set<String>> transfData = new HashMap<String,Set<String>>();
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(sourceFile));
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		int pIdx = 0;
		if(hasTitle){
			pIdx =1;
		}
		for (; pIdx <= rowCount; pIdx++) {
			XSSFRow row = sheet.getRow(pIdx);
			int columnCount = row.getLastCellNum();
			String name = row.getCell(0).getStringCellValue();
			for(int cIdx = 1;cIdx<columnCount;cIdx++){
				XSSFCell cell = row.getCell(cIdx);
				if(null != cell){
					String val = getExcelCellValue(cell).toUpperCase().trim();
					Set<String> nameSet = transfData.get(val);
					if(null == nameSet){
						nameSet = new HashSet<String>();
					}
					nameSet.add(name);
					transfData.put(val, nameSet);
				}				
			}
		}
		SXSSFWorkbook workbook2 = new SXSSFWorkbook(1000);
		SXSSFSheet sheet2 = workbook2.createSheet();
		int rowIdx = 0;
		for(String val:transfData.keySet()){
			SXSSFRow row = sheet2.createRow(rowIdx);
			row.createCell(0).setCellValue(val);
			Set<String> nameSet = transfData.get(val);
			int colId = 1;
			for(String name:nameSet){
				row.createCell(colId).setCellValue(name);
				colId ++;
			}
			rowIdx++;
		}
		writeSAPXlsx(workbook2, targetFile);
		System.out.println("转换成功,请查看文件["+targetFile.getAbsolutePath()+"]");
	}
	
	
	private static String getExcelCellValue(XSSFCell cell){
		if(cell.getCellTypeEnum() == CellType.NUMERIC){
			if(HSSFDateUtil.isCellDateFormatted(cell)){
				SimpleDateFormat df = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
				return df.format(cell.getDateCellValue());				
			}else{
				return ""+cell.getNumericCellValue();
			}
		}else if(cell.getCellTypeEnum() == CellType.FORMULA){
			return cell.getCellFormula();
		}else if(cell.getCellTypeEnum() == CellType.ERROR){
			return cell.getErrorCellString();
		}else if(cell.getCellTypeEnum() == CellType.BOOLEAN){
			return ""+cell.getBooleanCellValue();
		}else if(cell.getCellTypeEnum() == CellType.STRING){
			return cell.getStringCellValue();
		}else if(cell.getCellTypeEnum() == CellType.BLANK){
			return "";
		}else{
			return cell.getStringCellValue();
		}
	}
	private static void writeSAPXlsx(SXSSFWorkbook workBook, File targetFile) throws Throwable {
		OutputStream output = new FileOutputStream(targetFile);
		workBook.write(output);
		output.close();
	}
}
