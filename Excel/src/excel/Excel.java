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

        Sheet1 sheet1 = new Sheet1(workbook);
        sheet1.run();

        Sheet2 sheet2 = new Sheet2(workbook);
        sheet2.run();
        
        Sheet3 sheet3 = new Sheet3(workbook);
        sheet3.run();        
        
        try {
            FileOutputStream output = new FileOutputStream("/Users/kevinduraj/Desktop/kevin.xls");
            workbook.write(output);
            output.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
        
    }

}
