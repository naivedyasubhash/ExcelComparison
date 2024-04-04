package utility;

public class ExcelMain {

	public static void main(String[] args) {
		String path1="D:/Book1.xlsx";
		String path2="D:/Book2.xlsx";;
		ExcelComparison excel=new ExcelComparison(path1,path2);
		excel.ExcelCompare();

	}

}
