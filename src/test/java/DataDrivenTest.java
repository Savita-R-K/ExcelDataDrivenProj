import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;

public class DataDrivenTest {

    DataFormatter formatter=new DataFormatter();


    @Test(dataProvider = "getData")
    public void excelDataDriven(String greeting,String msg,String id){
        System.out.println(greeting+" "+msg+" "+id);
    }

    @DataProvider
    public Object[][] getData() throws IOException {
        FileInputStream fls=new FileInputStream("C:\\Users\\280679\\Intellij\\ExcelDataDrivenProj\\Book2.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fls);
        XSSFSheet sheet=workbook.getSheetAt(0);
        //no. of rows
        int rows=sheet.getPhysicalNumberOfRows();
        //no. of columns
        Row firstRow=sheet.getRow(0);
        int columns=firstRow.getLastCellNum();

        Object[][] data=new Object[rows-1][columns];
        for(int i=0;i<rows-1;i++){
            for(int j=0;j<columns;j++){
                data[i][j]=formatter.formatCellValue(sheet.getRow(i).getCell(j));
            }
        }

        return data;
    }

}
