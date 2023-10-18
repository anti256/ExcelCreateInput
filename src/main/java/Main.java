import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


public class Main {
    public static void main(String[] args) throws IOException {
        //создание объекта для связи с Excel
        XSSFWorkbook workbook = new XSSFWorkbook();

        //создание листа
        XSSFSheet spreadsheet = workbook.createSheet("MyExcelFile");

        //объект строка
        XSSFRow row;

        for (int i = 0; i < 10; i++) {//перебор строк
            row = spreadsheet.createRow(i);
            for (int j = 0; j < 10; j++) {//перебор колонок
                Cell cell = row.createCell(j);
                cell.setCellValue((i+1) + "/" + (j+1));
            }
        }

        //запись в файл
        File file = new File("F:\\myExcelFile.xlsx");
        FileOutputStream out = new FileOutputStream(file);

        workbook.write(out);
        out.close();


        row = spreadsheet.getRow(5);
        Cell cell = row.getCell(8);

        cell.getRow().setHeightInPoints(cell.getSheet().getDefaultRowHeightInPoints() * 2);

        CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.setWrapText(true);
        cell.setCellStyle(cellStyle);

        cell.setCellValue("изменение ячейки\nвторая строка");

        spreadsheet.autoSizeColumn(8);

        out = new FileOutputStream(file);
        workbook.write(out);
        out.close();
    }
}
