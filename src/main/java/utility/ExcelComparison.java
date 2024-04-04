package utility;

import java.io.FileInputStream;

//import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelComparison {
	String path1;
	String path2;
	FileInputStream file1;
	FileInputStream file2;
	XSSFWorkbook workbook1;
	XSSFWorkbook workbook2;
	XSSFSheet sheet1;
	XSSFSheet sheet2;
	XSSFRow row1;
	XSSFRow row2;
	XSSFCell cell1;
	XSSFCell cell2;

	ExcelComparison(String path1, String path2) {
		this.path1 = path1;
		this.path2 = path2;

	}

	public void ExcelCompare() {

		try {
			file1 = new FileInputStream(path1);
			file2 = new FileInputStream(path2);
			workbook1 = new XSSFWorkbook(file1);
			workbook2 = new XSSFWorkbook(file2);
			// find the number of sheets available
			int sheetNum1 = workbook1.getNumberOfSheets();
			int sheetNum2 = workbook2.getNumberOfSheets();
			if (sheetNum1 != sheetNum2) {
				System.out.println("File1 has :" + sheetNum1 + " number of Sheets.");
				System.out.println("File2 has :" + sheetNum2 + " number of Sheets.");
				return;
			}

			// Get first/desired sheet from the workbook
			for (int i = 0; i < sheetNum1; i++) {
				sheet1 = workbook1.getSheetAt(i);
				sheet2 = workbook2.getSheetAt(i);
				System.out.println("Comparing sheets :" + sheet1.getSheetName() + " with " + sheet2.getSheetName());
			}

			// Compare sheets
			if (compareTwoSheets(sheet1, sheet2)) {
				System.out.println("\n\nThe two excel sheets are Equal");

			} else {
				System.out.println("\n\nThe two excel sheets are Not Equal");

			}
			// close files
			file1.close();
			file2.close();

		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	/*
	 * This method is used to compare the sheets within excel sheets
	 */

	public boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2) {
		boolean equalSheets = true;
		int firstRow1 = sheet1.getFirstRowNum();
		int lastRow1 = sheet1.getLastRowNum();
		for (int i = firstRow1; i <= lastRow1; i++) {
			row1 = sheet1.getRow(i);
			row2 = sheet2.getRow(i);
			if (!compareTwoRows(row1, row2)) {
				equalSheets = false;
				System.out.println("Row " + i + " - Not Equal");
				break;
			} else {
				System.out.println("Row " + i + " - Equal");
			}
		}

		return equalSheets;

	}

	/*
	 * This method is used to compare the rows within excel sheets
	 */
	public boolean compareTwoRows(XSSFRow row1, XSSFRow row2) {

		int firstCell1 = row1.getFirstCellNum();
		int lastCell1 = row1.getLastCellNum();
		boolean equalRows = true;
		// Compare all cells in a row

		for (int i = firstCell1; i <= lastCell1; i++) {
			cell1 = row1.getCell(i);
			//System.out.println(cell1);
			cell2 = row2.getCell(i);
			//System.out.println(cell2);
			if (!compareTwoCells(cell1, cell2)) {
				equalRows = false;
				System.err.println("Cell " + i + " : Not Equal");
				System.out.println("Cell " + i + " of Downloaded Excel : " + cell1.getStringCellValue());
				System.out.println("Cell " + i + " of Baseline Excel : " + cell2.getStringCellValue());
			} else {
				System.out.println("Cell " + i + " - Equal");
				System.out.println("***********************************************************************");

			}
		}

		return equalRows;

	}

	/*
	 * This method is used to compare two cell values
	 */
	public boolean compareTwoCells(XSSFCell cell1,XSSFCell  cell2) {
		boolean equalcell=true;
		if(cell1!=null && cell2!=null ) {
			if(!cell1.toString().equalsIgnoreCase(cell2.toString())) {
				System.out.println(cell1.toString());
				System.out.println(cell2.toString());
				System.out.println("Difference Found at sheet");
				equalcell=false;
			}
		}
		return equalcell;
	}
	

}
