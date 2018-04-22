package cn.edu.xlxy.app.courses;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.stereotype.Service;

import cn.edu.xlxy.app.entity.CourseTable;

@Service
public class CoursesProcessor {

    public static void main(String[] args) throws Throwable {

	String fileName = "D:\\test\\CourseTable.xls";

	processCoursesTable(fileName);
    }

    public static List<CourseTable> processCoursesTable(String fileName) throws IOException {
	try {

	    FileInputStream fis = new FileInputStream(fileName);
	    HSSFWorkbook wb = new HSSFWorkbook(fis);

	    List<CourseTable> courseTables = new ArrayList<CourseTable>();

	    for (int k = 0; k < wb.getNumberOfSheets(); k++) {
		HSSFSheet sheet = wb.getSheetAt(k);
		int rows = sheet.getPhysicalNumberOfRows();
		// System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " +
		// rows + " row(s).");
		/**
		 * for (int r = 0; r < rows; r++) { HSSFRow row = sheet.getRow(6); if (row ==
		 * null) { continue; }
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
		String classroomName = "";

		System.out.println("\nROW " + row.getRowNum() + " has " + row.getPhysicalNumberOfCells() + " cell(s).");
		for (int c = 0; c < row.getLastCellNum(); c++) {
		    HSSFCell cell = row.getCell(c);
		    String value;
		    String courseName = "";
		    String weekSequence = "";
		    String section = "";
		    int startWeek = 0;
		    int endWeek = 0;
		    String weekType = "";
		    String studentInformation = "";

		    if (cell != null && cell.getCellTypeEnum().equals(CellType.STRING)) {

			String cellContent = cell.getStringCellValue();

			// System.out.println("current cell content is " + cellContent);
			// System.out.println("current cell number is " + c);
			String[] splitedString = cellContent.split("\\[");
			if (splitedString.length == 1) {
			    // it must be the classroom name ;
			    classroomName = cellContent;

			} else if (splitedString.length == 2) {

			    courseName = splitedString[0];
			    String contentLeft = splitedString[1];
			    splitedString = contentLeft.split("\\]");
			    if (splitedString.length > 1) {

				weekSequence = splitedString[0];

				if (weekSequence.contains("单周")) {
				    weekType = "even";
				} else if (weekSequence.contains("双周")) {
				    weekType = "odd";
				} else {
				    weekType = "normal";
				}
				String weekInfo = weekSequence.replace("单周", "").replace("双周", "").replace("周", "")
					.trim();
				String[] splitedWeekInfo = weekInfo.split("-");
				startWeek = Integer.valueOf(splitedWeekInfo[0]);
				endWeek = Integer.valueOf(splitedWeekInfo[1]);

				contentLeft = splitedString[1];

				splitedString = contentLeft.split("；");
				if (splitedString.length > 1) {
				    section = splitedString[0];
				    contentLeft = splitedString[1];
				    studentInformation = contentLeft;
				}
			    }

			    System.out.print("course info: " + courseName + " ");
			    System.out.print("weeks info: " + weekSequence + " ");
			    System.out.print("week info: " + (c - 1) / 5 + " ");
			    System.out.print("section info: " + section + " ");
			    System.out.println("student info: " + contentLeft + " ");
			} else if (splitedString.length > 2) {

			}

			switch (weekType) {
			case "odd":
			    for (int i = startWeek + startWeek % 2; i < endWeek; i++) {
				CourseTable courseTable = new CourseTable();

				courseTable.setClassroomName(classroomName);
				courseTable.setCourseInformation(courseName);
				courseTable.setWeekSequence(String.valueOf(i));
				courseTable.setAvailable("FALSE");
				courseTable.setWeekday(String.valueOf((c - 1) / 5 + 1));
				courseTable.setSection(section);
				courseTable.setStudentInformation(studentInformation);

				courseTables.add(courseTable);
			    }

			    break;
			case "even":
			    for (int i = startWeek + (startWeek - 1) % 2; i < endWeek; i += 2) {
				CourseTable courseTable = new CourseTable();

				courseTable.setClassroomName(classroomName);
				courseTable.setCourseInformation(courseName);
				courseTable.setWeekSequence(String.valueOf(i));
				courseTable.setAvailable("FALSE");
				courseTable.setWeekday(String.valueOf((c - 1) / 5 + 1));
				courseTable.setSection(section);
				courseTable.setStudentInformation(studentInformation);

				courseTables.add(courseTable);
			    }
			    break;

			default:
			    for (int i = startWeek; i < endWeek; i++) {
				CourseTable courseTable = new CourseTable();

				courseTable.setClassroomName(classroomName);
				courseTable.setCourseInformation(courseName);
				courseTable.setWeekSequence(String.valueOf(i));
				courseTable.setAvailable("FALSE");
				courseTable.setWeekday(String.valueOf((c - 1) / 5 + 1));
				courseTable.setSection(section);
				courseTable.setStudentInformation(studentInformation);

				courseTables.add(courseTable);
			    }

			    break;

			}

		    }
		}
	    }
	    // }
	    wb.close();
	    return courseTables;

	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	}
	return null;
    }

}
