import java.io.IOException;
import java.util.ArrayList;

public class SampleTestCase{
    public static void main(String[] args) throws IOException {
        ExcelData dataDriven=new ExcelData();
        ArrayList<String> testData=dataDriven.getExcelData("InvalidLogin");
        for(String data:testData){
            System.out.println(data);
        }
    }
}
