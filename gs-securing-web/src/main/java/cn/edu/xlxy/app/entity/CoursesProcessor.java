package cn.edu.xlxy.app.entity;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CoursesProcessor {

    public static void main(String[] args) throws Throwable {

	String fileName = "D:\\test\\WholeSchoolCourseTable1.xls";
	
	processCoursesTable(fileName);
    }

    public static void processCoursesTable(String fileName) throws IOException {
	try {
	    FileInputStream fis = new FileInputStream(fileName);
	    HSSFWorkbook wb = new HSSFWorkbook(fis);

	    for (int k = 0; k < wb.getNumberOfSheets(); k++) {
		HSSFSheet sheet = wb.getSheetAt(k);
		int rows = sheet.getPhysicalNumberOfRows();
		System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows + " row(s).");
		for (int r = 0; r < rows; r++) {
		    HSSFRow row = sheet.getRow(r);
		    if (row == null) {
			continue;
		    }

		    System.out.println(
			    "\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
		    for (int c = 0; c < row.getLastCellNum(); c++) {
			HSSFCell cell = row.getCell(c);
			String value;

			if (cell != null) {
			    switch (cell.getCellTypeEnum()) {

			    case FORMULA:
				value = "FORMULA value=" + cell.getCellFormula();
				break;

			    case NUMERIC:
				value = "NUMERIC value=" + cell.getNumericCellValue();
				break;

			    case STRING:
				value = "STRING value=" + cell.getStringCellValue();
				break;

			    case BLANK:
				value = "<BLANK>";
				break;

			    case BOOLEAN:
				value = "BOOLEAN value-" + cell.getBooleanCellValue();
				break;

			    case ERROR:
				value = "ERROR value=" + cell.getErrorCellValue();
				break;

			    default:
				value = "UNKNOWN value of type " + cell.getCellTypeEnum();
			    }
			    System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE=" + value);
			}
		    }
		}
	    }
	    wb.close();

	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	}
    }
    public static void processCoursesTable1(String fileName) throws IOException {
	try {
	    FileInputStream fis = new FileInputStream(fileName);
	    HSSFWorkbook wb = new HSSFWorkbook(fis);
	    
	    for (int k = 0; k < wb.getNumberOfSheets(); k++) {
		HSSFSheet sheet = wb.getSheetAt(k);
		int rows = sheet.getPhysicalNumberOfRows();
		System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows + " row(s).");
		for (int r = 0; r < rows; r++) {
		    HSSFRow row = sheet.getRow(r);
		    if (row == null) {
			continue;
		    }
		    
		    System.out.println(
			    "\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
		    for (int c = 0; c < row.getLastCellNum(); c++) {
			HSSFCell cell = row.getCell(c);
			String value;
			
			if (cell != null) {
			    switch (cell.getCellTypeEnum()) {
			    
			    case FORMULA:
				value = "FORMULA value=" + cell.getCellFormula();
				break;
				
			    case NUMERIC:
				value = "NUMERIC value=" + cell.getNumericCellValue();
				break;
				
			    case STRING:
				value = "STRING value=" + cell.getStringCellValue();
				break;
				
			    case BLANK:
				value = "<BLANK>";
				break;
				
			    case BOOLEAN:
				value = "BOOLEAN value-" + cell.getBooleanCellValue();
				break;
				
			    case ERROR:
				value = "ERROR value=" + cell.getErrorCellValue();
				break;
				
			    default:
				value = "UNKNOWN value of type " + cell.getCellTypeEnum();
			    }
			    System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE=" + value);
			}
		    }
		}
	    }
	    wb.close();
	    
	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	}
    }

}
