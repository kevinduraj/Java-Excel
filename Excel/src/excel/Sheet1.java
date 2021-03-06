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

        Sheet sheet = workbook.createSheet("Stat-1");

        List<Cell> cells = new ArrayList<>();

        for (int i = 0; i < 10; i++) {
            for (int j = 0; j < 10; j++) {
                cell = sheet.createRow(i).createCell(j);
                cells.add(cell);
            }
        }

        for (Cell one : cells) {
            one.setCellValue(1 + (int) (Math.random() * 100));
        }

    }
}
