package excel;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

class Sheet2 {

    Workbook workbook;
    Sheet sheet;

    public Sheet2(Workbook workbook) {
        this.workbook = workbook;
        sheet = workbook.createSheet("Stat-2");
    }

    public void run() {

        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);

            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(1 + (int) (Math.random() * 100));
            }
        }
    }

}
