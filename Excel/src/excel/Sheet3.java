package excel;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/*----------------------------------------------------------------------------*/
class Sheet3 {

    Workbook workbook;
    Sheet sheet;
    Row[] rows;
    Cell[] cells;
    int width = 100;
    int height = 100;

    /*------------------------------------------------------------------------*/
    public Sheet3(Workbook workbook) {
        this.workbook = workbook;
        sheet = workbook.createSheet("Stat-3");
        rows = new Row[width];
        cells = new Cell[height];
        initialize();
    }
    /*------------------------------------------------------------------------*/

    public final void initialize() {

        for (int i = 0; i < width; i++) {
            rows[i] = sheet.createRow(i);

            for (int j = 0; j < height; j++) {
                cells[j] = rows[i].createCell(j);
            }
        }
    }
    /*------------------------------------------------------------------------*/

    public void randomValues() {

        for (int i = 0; i < 10; i++) {
            for (int j = 0; j < 10; j++) {
                cellValue(i, j, (1 + (int) (Math.random() * 100)));
            }
        }

        cellValue(0, 0, 0);
        cellValue(10, 10, 10);
    }
    /*------------------------------------------------------------------------*/

    public void cellValue(int i, int j, int value) {

        Cell cell = rows[i].getCell(j);
        cell.setCellValue(value);

    }


    /*------------------------------------------------------------------------*/
}
/*---------------------------------Sheet3-------------------------------------*/