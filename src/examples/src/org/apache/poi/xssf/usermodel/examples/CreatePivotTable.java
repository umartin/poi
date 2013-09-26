package org.apache.poi.xssf.usermodel.examples;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 */
public class CreatePivotTable {
    
    public static void main(String[] args) throws FileNotFoundException, IOException {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();
        sheet.createPivotTable();
  
        FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.xlsx");
        wb.write(fileOut);
        fileOut.close(); 
    }

}
