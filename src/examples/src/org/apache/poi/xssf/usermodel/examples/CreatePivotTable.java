package org.apache.poi.xssf.usermodel.examples;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
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
/*
        wb.addPivotCache();
        wb.addPivotCache();
        
       */ 
        FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.xlsx");
        wb.write(fileOut);
        fileOut.close(); 
    }

}
