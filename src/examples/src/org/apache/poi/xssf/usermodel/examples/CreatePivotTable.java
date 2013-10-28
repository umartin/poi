/* ====================================================================
 Licensed to the Apache Software Foundation (ASF) under one or more
 contributor license agreements.  See the NOTICE file distributed with
 this work for additional information regarding copyright ownership.
 The ASF licenses this file to You under the Apache License, Version 2.0
 (the "License"); you may not use this file except in compliance with
 the License.  You may obtain a copy of the License at

 http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 ==================================================================== */
package org.apache.poi.xssf.usermodel.examples;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSourceType;

/**
 *
 */
public class CreatePivotTable {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        createPivot("ooxml-pivottable.xlsx");
        createPivot("ooxml-pivottable.zip");
    }
    
    public static void createPivot(String fileName)throws FileNotFoundException, IOException{
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();
        setCellData(sheet);
        XSSFPivotTable pivotTable = sheet.createPivotTable(new AreaReference("A1:C3"), "H5");
                
        CTPivotTableDefinition definition = pivotTable.getCTPivotTableDefinition();
        //Set later
        definition.setMultipleFieldFilters(false);
        //Look up definition
        definition.setOutline(true);
        definition.setOutlineData(true);
        
        int rowStart = 0;
        int rowEnd = 2;
        
        int columnStart = 0;
        int columnEnd = 2;
        
        pivotTable.setReferences(columnStart, rowStart, columnEnd, rowEnd);
        pivotTable.createCacheRecords();
   
        CTPivotCacheDefinition cacheDef = pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition();
        cacheDef.setRecordCount(pivotTable.getPivotCacheRecords().getCtPivotCacheRecords().getCount());
        
        CTCacheFields cFields = cacheDef.addNewCacheFields();
        //Create cachefield/s and empty SharedItems
        pivotTable.getPivotCacheDefinition().createCacheFields(sheet);
                       
        pivotTable.addRowLabel(0);
        //pivotTable.addColumnLabel(STDataConsolidateFunction.AVERAGE,1);
        pivotTable.addColumnLabel(STDataConsolidateFunction.SUM, 2);
        pivotTable.addRowLabel(1);
        
        FileOutputStream fileOut = new FileOutputStream(fileName);
        wb.write(fileOut);
        fileOut.close();
    }
    
    public static void setCellData(XSSFSheet sheet){
        Row row1 = sheet.createRow((short) 0);
        // Create a cell and put a value in it.
        Cell cell = row1.createCell((short) 0);
        cell.setCellValue("Names");
        Cell cell2 = row1.createCell((short) 1);
        cell2.setCellValue("#");
        Cell cell7 = row1.createCell((short) 2);
        cell7.setCellValue("%");

        Row row2 = sheet.createRow((short) 1);
        Cell cell3 = row2.createCell((short) 0);
        cell3.setCellValue("Jessica");
        Cell cell4 = row2.createCell((short) 1);
        cell4.setCellValue(3);
        Cell cell8 = row2.createCell((short) 2);
        cell8.setCellValue(85);

        Row row3 = sheet.createRow((short) 2);
        Cell cell5 = row3.createCell((short) 0);
        cell5.setCellValue("Sofia");
        Cell cell6 = row3.createCell((short) 1);
        cell6.setCellValue(3);
        Cell cell9 = row3.createCell((short) 2);
        cell9.setCellValue(82);
        
        //Under pivottable
        Row row6 = sheet.createRow((short) 5);
        Cell cell10 = row6.createCell((short) 5);
        cell10.setCellValue("Hej");
        
    }
}