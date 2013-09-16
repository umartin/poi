/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.apache.poi.xssf.usermodel.examples;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Demonstrates how to work with excel tables.
 *
 * @author Sofia Larsson
 */
public class CreateTable {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
     
        FileOutputStream fileOut = new FileOutputStream("ooxml-table.xlsx");
        
        wb.write(fileOut);
        fileOut.close();
    }
}
