package cn.edu.xlxy.app.courses;

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
		/**
		 * for (int r = 0; r < rows; r++) { HSSFRow row =
		 * sheet.getRow(6); if (row == null) { continue; }
		 */
		HSSFRow row = sheet.getRow(6);
		if (row == null) {
		    continue;
		}
		// Monday: Column 2 - 6
		// Tuesday: Column 7 - 11
		// Wednesday: Column 12 - 16
		// Thursday: Column 17 - 21
		// Friday: Column 22 - 16
		// Saturday: Column 27 - 31
		// Sunday: Column 32 - 36

		System.out.println("\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
		for (int c = 0; c < row.getLastCellNum(); c++) {
		    HSSFCell cell = row.getCell(c);
		    String value;
		    String courseName = "";
		    String weekNumber = "";
		    String section = "";

		    if (cell != null) {
			switch (cell.getCellTypeEnum()) {

			case FORMULA:
			    value = "FORMULA value=" + cell.getCellFormula();
			    break;

			case NUMERIC:
			    value = "NUMERIC value=" + cell.getNumericCellValue();
			    break;

			case STRING:
			    // value = "STRING value=" +
			    // cell.getStringCellValue();

			    String cellContent = cell.getStringCellValue();
			    String[] splitedString = cellContent.split("\\[");
			    if (splitedString.length > 1) {

				courseName = splitedString[0];
				String leftPart = splitedString[1];
				splitedString = leftPart.split("\\]");
				if (splitedString.length > 1) {

				    weekNumber = splitedString[0];
				    leftPart = splitedString[1];

				    splitedString = leftPart.split("£»");
				    if (splitedString.length > 1) {
					section = splitedString[0];
					leftPart = splitedString[1];
				    }
				}

				System.out.print("course info: " + courseName + " ");
				System.out.print("weeks info: " + weekNumber + " ");
				System.out.print("week info: " + (c-1)/5 + " ");
				System.out.print("section info: " + section + " ");
				System.out.println("student info: " + leftPart + " ");
			    }

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
			// System.out.println("CELL col=" +
			// cell.getColumnIndex() + " VALUE=" + value);
		    }
		}
	    }
	    // }
	    wb.close();

	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	}
    }

}
