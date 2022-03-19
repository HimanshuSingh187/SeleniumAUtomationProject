package SeleniumJava;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class demoExcelDriven {
	
	public ArrayList<String> excelDriven(String testCaseName) throws IOException{
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(
				"C:\\Users\\himanshu.singhs\\practiseprojects\\MavenProject\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int numberofSheets = workbook.getNumberOfSheets();

		for (int i = 0; i < numberofSheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Login")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.rowIterator();
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator();
				int k = 0;
				int column = 0;
				while (cell.hasNext()) {
					Cell cellValue = cell.next();
					if (cellValue.getStringCellValue().equalsIgnoreCase("TestCases")) {
						System.out.println(cellValue.getStringCellValue());
						System.out.println("I am here");
						column = k;
						// System.out.println(column);
					}
					k++;
				}
				System.out.println(column);

				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						Iterator<Cell> c = r.cellIterator();

						while (c.hasNext()) {
							Cell cv = c.next();
							if(cv.getCellType()==CellType.STRING){
								a.add(cv.getStringCellValue());
							}
							else{
								a.add(NumberToTextConverter.toText(cv.getNumericCellValue()));
							}
						}
					}
				}

			}
		}
		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
	}

}
