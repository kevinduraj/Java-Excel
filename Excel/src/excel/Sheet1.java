package excel;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

class Sheet1 {

    Workbook workbook;
    Cell cell;

    public Sheet1(Workbook workbook) {
        this.workbook = workbook;
    }

    public void run() {
        //Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Statistics");

        List<Cell> cells = new ArrayList<>();

        for (int i = 0; i < 20; i++) {
            for (int j = 0; j < 20; j++) {
                cell = sheet.createRow(i).createCell(j);
                cells.add(cell);
            }
            System.out.print("\n" + i);
        }

        for (Cell one : cells) {
            one.setCellValue(1 + (int) (Math.random() * 100));
        }

        try {
            FileOutputStream output = new FileOutputStream("/Users/kevinduraj/Desktop/kevin.xls");
            workbook.write(output);
            output.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
