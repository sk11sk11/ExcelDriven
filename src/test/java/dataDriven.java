import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public static void main(String[] args) throws IOException {

		
		
		

		// open a channel to read data:
		FileInputStream fis = new FileInputStream("C:\\Users\\C-T74A897\\Documents\\demodata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// Identify Testcases col by scanning the entire 1st row.

				Iterator<Row> rows = sheet.iterator(); // sheet is a collection of rows
				Row firstrow = rows.next(); // access first row
				Iterator<Cell> ce = firstrow.cellIterator(); // row is a collection of cells

				int k = 0;
				int column = 0;

				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {

						column = k;
					}
					k++;
				}
				System.out.println(column);

				// Once Col is identified, then scan entire testcase Col to identify "Purchase" testcase row.
				
				while(rows.hasNext()) {
					
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
						
						// after you grab "Purchase" testcase row, pull all the data and feed into test.
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							System.out.println(cv.next().getStringCellValue());
						}
						
					}
					
				}
				
				
			}

		}

	}

}
