import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelData {

    public ArrayList<String> getExcelData(String TestCaseName) throws IOException {
        FileInputStream file = new FileInputStream("C:\\Users\\savita\\IdeaProjects\\ExcelDataDriven\\Book2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        int sheets = workbook.getNumberOfSheets();
        ArrayList<String> a = new ArrayList<>();
        for (int i = 0; i < sheets; i++) {
            //--------------identifying sheet----------------------------------
            if (workbook.getSheetName(i).equalsIgnoreCase("Testdata")) {
                //required sheet
                XSSFSheet sheet = workbook.getSheetAt(0);
                //rows in the sheet
                Iterator<Row> rows = sheet.iterator();
                //1st row->column headings to move to required column
                Row firstRow = rows.next();
                //iterating across column headings ->each cell of first row
                Iterator<Cell> cell = firstRow.cellIterator();
                //to grab the column no. of TestCases.
                int k = 0;
                int column = 0;
                while (cell.hasNext()) {
                    Cell value = cell.next();
                    //----------------------------------Identifying column heading---------------------------
                    if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
                        column = k;
                    }
                    k++;
                }

                //--------------------------Identifying the required column data - required row data--------------------------------
                while (rows.hasNext()) {
                    Row row = rows.next();
                    if (row.getCell(column).getStringCellValue().equalsIgnoreCase(TestCaseName)) {
                        Iterator<Cell> val = row.cellIterator();
                        while (val.hasNext()) {
                            Cell c=val.next();
                            if(c.getCellType()== CellType.STRING){
                                a.add(c.getStringCellValue());
                            }else {
                                a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
                            }
                        }
                    }
                }
            }
        }
        return a;
    }

}
