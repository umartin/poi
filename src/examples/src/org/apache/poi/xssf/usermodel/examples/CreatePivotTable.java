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
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSharedItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSourceType;

/**
 *
 */
public class CreatePivotTable {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();
        XSSFPivotTable pivotTable = sheet.createPivotTable();

        Row row1 = sheet.createRow((short) 0);
        // Create a cell and put a value in it.
        Cell cell = row1.createCell((short) 0);
        cell.setCellValue("Names");
        Cell cell2 = row1.createCell((short) 1);
        cell2.setCellValue("#");

        Row row2 = sheet.createRow((short) 1);
        Cell cell3 = row2.createCell((short) 0);
        cell3.setCellValue("Jessica");
        Cell cell4 = row2.createCell((short) 1);
        cell4.setCellValue(3);

        Row row3 = sheet.createRow((short) 2);
        Cell cell5 = row3.createCell((short) 0);
        cell5.setCellValue("Sofia");
        Cell cell6 = row3.createCell((short) 1);
        cell6.setCellValue(3);
        
        CTPivotTableDefinition definition = pivotTable.getCTPivotTableDefinition();
        definition.setMultipleFieldFilters(false);
        definition.setOutline(true);
        definition.setOutlineData(true);
        definition.setDataCaption("Values");
        CTLocation location = definition.addNewLocation();
        location.setFirstDataCol(1);
        location.setFirstDataRow(1);
        location.setFirstHeaderRow(1);
        location.setRef("F5:H22");
        CTPivotFields fields = definition.addNewPivotFields();
        fields.setCount(2);
        CTPivotField field = fields.addNewPivotField();
        field.setShowAll(false);
        CTPivotField field2 = fields.addNewPivotField();
        field2.setShowAll(false);

        //Cache definition
        CTPivotCacheDefinition cacheDef = pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition();
        CTCacheSource source = cacheDef.addNewCacheSource();
        source.setType(STSourceType.WORKSHEET);
        CTWorksheetSource worksheetSource = source.addNewWorksheetSource();
        worksheetSource.setSheet(sheet.getSheetName());
        worksheetSource.setRef("A1:B3");
        
        CTCacheFields cFields = cacheDef.addNewCacheFields();
        cFields.setCount(2);
        

        int rowStart = 0;
        int rowEnd = 2;
        
        int columnStart = 0;
        int columnEnd = 1;
        
        pivotTable.setReferences(columnStart, rowStart, columnEnd, rowEnd);
        pivotTable.createCacheRecords();
   
        cacheDef.setRecordCount((rowEnd-rowStart) +1);
        Row row;
                
        for(int i=columnStart; i<=columnEnd; i++) {
            row = sheet.getRow(rowStart);
            CTCacheField cf = cFields.addNewCacheField();
            //General number format
            cf.setNumFmtId(0);
            cf.setName(row.getCell(i).getStringCellValue());
            CTSharedItems shared = cf.addNewSharedItems();
            shared.setCount(0);
            shared.setContainsBlank(false);
            shared.setContainsDate(false);
            shared.setContainsInteger(false);
            shared.setContainsMixedTypes(false);
            shared.setContainsNonDate(false);
            shared.setContainsNumber(false);
            shared.setContainsSemiMixedTypes(false);
            shared.setContainsString(false);
            /*for(int j=rowStart+1; j<=rowEnd; j++) {
                row = sheet.getRow(j);
                c = row.getCell(i);
                switch (c.getCellType()) {
                    case (Cell.CELL_TYPE_BOOLEAN):
                        shared.addNewB().setV(c.getBooleanCellValue());
                        shared.setCount(shared.getCount()+1);
                        break;
                    case (Cell.CELL_TYPE_STRING):
                        shared.addNewS().setV(c.getStringCellValue());
                        shared.setCount(shared.getCount()+1);
                        shared.setContainsSemiMixedTypes(true);
                        shared.setContainsString(true);
                        break;
                    case (Cell.CELL_TYPE_NUMERIC):
                        shared.addNewN().setV(c.getNumericCellValue());
                        shared.setCount(shared.getCount()+1);
                        shared.setContainsNumber(true);
                        if(c.getNumericCellValue()%1==0) {
                            shared.setContainsInteger(true);
                        }
                        if(HSSFDateUtil.isCellDateFormatted(c)) {
                            shared.setContainsDate(true);
                        }
                        break;
                    case (Cell.CELL_TYPE_BLANK):
                        shared.addNewM();
                        shared.setContainsBlank(true);
                        shared.setCount(shared.getCount()+1);
                        break;
                    case (Cell.CELL_TYPE_ERROR):
                        shared.addNewE().setV(String.valueOf(c.getErrorCellValue()));
                        shared.setCount(shared.getCount()+1);
                        break;
                    case (Cell.CELL_TYPE_FORMULA):
                        shared.addNewS().setV(c.getCellFormula());
                        shared.setCount(shared.getCount()+1);
                        break;
                    default:
                        
                        break;
                }
            }
            if(shared.getBList().size() > 0 && shared.getBList().size() < shared.getCount() ||
                    shared.getDList().size() > 0 && shared.getDList().size() < shared.getCount() ||
                    shared.getEList().size() > 0 && shared.getEList().size() < shared.getCount() ||
                    shared.getMList().size() > 0 && shared.getMList().size() < shared.getCount() ||
                    shared.getNList().size() > 0 && shared.getNList().size() < shared.getCount() ||
                    shared.getSList().size() > 0 && shared.getSList().size() < shared.getCount()) {
                shared.setContainsMixedTypes(true);
            }
            if(shared.getDList().size() < shared.getCount()) {
                shared.setContainsNonDate(true);
            }
            if(shared.getContainsNumber()){
                Iterator<CTNumber> it = shared.getNList().iterator();
                Double max = Double.MIN_VALUE;
                Double value;
                while(it.hasNext()) {
                    value = it.next().getV();
                    if(value > max) {
                        max = value;
                    } 
                }
                shared.setMaxValue(max);
                it = shared.getNList().iterator();
                Double min = Double.MAX_VALUE;
                while(it.hasNext()) {
                    value = it.next().getV();
                    if(value < min) {
                        min = value;
                    } 
                }
                shared.setMinValue(min);
            }
*/
        }
        /*Row columnRow = sheet.getRow((short) 0);
            for (int i = 0; i < columnRow.getLastCellNum(); i++) {
                CTCacheField cf = cFields.addNewCacheField();
                cf.setNumFmtId(0);
                cf.setName(columnRow.getCell(i).getStringCellValue());
                Iterator iterator = sheet.rowIterator();
                it.next();
                while (it.hasNext()) {
                    CTSharedItems shared = cf.addNewSharedItems();
                }
            }*/
        FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }
}