/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.apache.poi.xssf.usermodel.examples;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Jessica
 */
public class TestCreatingDifferentSheets {
    
    private static String[] sales = {"$41,107", "$72,707", "$41,676", "$87,858", "$45,606", "$49,017", "$57,967", "$70,702", "$77,738", "$69,496"};
    private static Integer[] orders = {217, 168, 224, 286, 226, 228, 234, 267, 279, 261};
    private static String[] regions = {"West", "West", "North", "North", "South", "East", "West", "East", "East", "South"};
    private static String[] names = {"Bill", "Frank", "Harry", "Janet", "Joe", "Martha", "Mary", "Ralph", "Sam", "Tom"};
    private static String[] header = {"SalesRep", "Region", "# Orders", "Total Sales"};
    private static String titel = "Coockie Sales by Region";
    
    public static void main(String[] args) throws FileNotFoundException, IOException{
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet");
        
        for(int i = 0; i < 12; i++){            
            Row row = sheet.createRow(i);
            for(int k = 0; k < 4; k++){
                Cell c = row.createCell(k);
                if(i == 0 && k == 0){
                    c.setCellValue(titel);
                }else if(i == 1){
                    c.setCellValue(header[k]);
                }else if( i > 1 && k == 0){
                    c.setCellValue(names[i-2]);
                }else if( i > 1 && k == 1){
                    c.setCellValue(regions[i-2]);
                }else if( i > 1 && k == 2){
                    c.setCellValue(orders[i-2]);
                }else if(i > 1 && k == 3){
                    c.setCellValue(sales[i-2]);
                }
            }            
        }        
        
        FileOutputStream out =  new FileOutputStream("test.xlsx");
        wb.write(out);
        out.close();
    }
        
}
