import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
public class Main {
    public static void main(String[] args) {
        //создание объекта для связи с Excel
        XSSFWorkbook workbook = new XSSFWorkbook();

        //создание листа
        XSSFSheet spreadsheet = workbook.createSheet("MyExcelFile");

        //объект строка
        XSSFRow row;

        //счетчик строк
        int rowid = 0;

        for (int i = 0; i < 10; i++) {
            row = spreadsheet.createRow(rowid++);
            //int cellid = 0;//счетчик колонок
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(i + "/" + j);
            }
        }

        FileOutputStream out = new FileOutputStream(
                new File("C:/savedexcel/myExcelFile.xlsx"));

        workbook.write(out);
        out.close();
    }
}
