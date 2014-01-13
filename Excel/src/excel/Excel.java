/*
 * Building Excel Spreadsheet using Apache2 POI
 * http://poi.apache.org/download.html
 * Written by: Kevin T. Duraj
 */
package excel;

import java.io.FileNotFoundException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Excel {

    public static void main(String[] args) {

        Workbook workbook = new HSSFWorkbook();

        Sheet1 static1 = new Sheet1(workbook);
        static1.run();

    }

}
