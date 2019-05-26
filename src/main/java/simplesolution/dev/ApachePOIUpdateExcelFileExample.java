package simplesolution.dev;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.InputStream;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePOIUpdateExcelFileExample {

    public static void main(String... args) {
        try(InputStream inputStream = new FileInputStream("sample.xlsx")) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            Cell cell = row.getCell(1);
            if(cell == null) {
                cell = row.createCell(1);
            }

            cell.setCellValue("SimpleSolution.dev");

            try(OutputStream outputStream = new FileOutputStream("sample.xlsx")) {
                workbook.write(outputStream);
            }
        }catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
