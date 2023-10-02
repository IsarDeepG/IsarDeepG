package com.excel.utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MyXLSReader {
	
	public String filepath;
	FileInputStream fis=null;;
	Workbook workbook=null;;
	Sheet sheet=null;;
	Row row=null;;
	Cell cell=null;;
	public  FileOutputStream fileOut =null;
	String fileExtension=null;
		
	public MyXLSReader(String filepath) throws IOException{
		
		this.filepath = filepath;
		fileExtension = filepath.substring(filepath.indexOf(".x"));
		
	   try {
			fis = new FileInputStream(filepath);
			
			if(fileExtension.equals(".xlsx")){
				
				workbook = new XSSFWorkbook(fis);
				
				
			} else if(fileExtension.equals(".xls")){
				
				workbook = new HSSFWorkbook(fis);
				
			}
			
			sheet = workbook.getSheetAt(0);	
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			fis.close();			
		}
		
	}
	
	// returns the row count in a sheet
	public int getRowCount(String sheetname){
		
		int sheetIndex = workbook.getSheetIndex(sheetname);
		if(sheetIndex==-1){			
			return 0;
		} else {			
			sheet = workbook.getSheetAt(sheetIndex);
			int rowsCount = sheet.getLastRowNum()+1;
			return rowsCount;		
		}
		
	}	
	
	
	// returns the data from a cell
	public String getCellData(String sheetname, String colName, int rowNum) {
	    try {
	        if (rowNum <= 0)
	            return "";

	        int sheetIndex = workbook.getSheetIndex(sheetname);
	        if (sheetIndex == -1)
	            return "";

	        sheet = workbook.getSheetAt(sheetIndex);
	        row = sheet.getRow(0);
	        int colNum = -1;

	        for (int i = 0; i < row.getLastCellNum(); i++) {
	            if (row.getCell(i).getStringCellValue().equals(colName))
	                colNum = i;
	        }

	        if (colNum == -1)
	            return "";

	        sheet = workbook.getSheetAt(sheetIndex);
	        row = sheet.getRow(rowNum - 1);
	        if (row == null)
	            return "";

	        cell = row.getCell(colNum);
	        if (cell == null)
	            return "";

	        switch (cell.getCellType()) {
	            case NUMERIC:
	                if (DateUtil.isCellDateFormatted(cell)) {
	                    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
	                    return sdf.format(cell.getDateCellValue());
	                } else {
	                    return String.valueOf(cell.getNumericCellValue());
	                }
	            case BOOLEAN:
	                return String.valueOf(cell.getBooleanCellValue());
	            case STRING:
	                return cell.getStringCellValue();
	            case BLANK:
	                return "";
	            default:
	                return cell.getStringCellValue();
	        }
	    } catch (Exception e) {
	        e.printStackTrace();
	        return "row " + rowNum + " or column " + colName + " does not exist in xls";
	    }
	}
	
	// returns the data from a cell
	public String getCellData(String sheetname, int colNum, int rowNum) {
	    try {
	        if (rowNum <= 0)
	            return "";

	        int sheetIndex = workbook.getSheetIndex(sheetname);

	        if (sheetIndex == -1)
	            return "";

	        sheet = workbook.getSheetAt(sheetIndex);
	        row = sheet.getRow(rowNum - 1);
	        if (row == null)
	            return "";
	        cell = row.getCell(colNum - 1);
	        if (cell == null)
	            return "";

	        switch (cell.getCellType()) {
	            case STRING:
	                return cell.getStringCellValue();
	            case NUMERIC:
	                if (DateUtil.isCellDateFormatted(cell)) {
	                    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
	                    return sdf.format(cell.getDateCellValue());
	                } else {
	                    return String.valueOf(cell.getNumericCellValue());
	                }
	            case BOOLEAN:
	                return String.valueOf(cell.getBooleanCellValue());
	            case BLANK:
	                return "";
	            default:
	                return cell.getStringCellValue();
	        }
	    } catch (Exception e) {
	        e.printStackTrace();
	        return "row " + rowNum + " or column " + colNum + " does not exist in xls";
	    }
	}
	
		
}