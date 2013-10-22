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
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTI;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSharedItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STAxis;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;
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
        XSSFPivotTable pivotTable = sheet.createPivotTable();

        setCellData(sheet);
                
        CTPivotTableDefinition definition = pivotTable.getCTPivotTableDefinition();
        //Set later
        definition.setMultipleFieldFilters(false);
        //Look up definition
        definition.setOutline(true);
        definition.setOutlineData(true);
        
        pivotTable.setLocation("F5:G8", 1, 1, 1);
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
   
        cacheDef.setRecordCount(pivotTable.getPivotCacheRecords().getCtPivotCacheRecords().getCount());
        Row row;
        Cell c;     
        //kolla om det spelar nÃ¥gon roll att vi alltid sÃ¤tter?
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
            for(int j=rowStart+1; j<=rowEnd; j++) {
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
                        //shared.addNewN().setV(c.getNumericCellValue());
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
            /*if(shared.getContainsNumber()){
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
            }  */
        }
       
        //Set dynamically when knowing how many fields to display
        CTPivotFields fields = definition.addNewPivotFields();
        fields.setCount(2);
        CTPivotField field = fields.addNewPivotField();
        field.setShowAll(true);
        field.setAxis(STAxis.AXIS_ROW);
        CTPivotField field2 = fields.addNewPivotField();
        field2.setShowAll(true);
        field2.setDataField(true);
        
        //Fill out the fields
        CTItems items = field.addNewItems();
        items.addNewItem().setT(STItemType.DEFAULT);
        items.addNewItem().setT(STItemType.DEFAULT);
        items.addNewItem().setT(STItemType.DEFAULT);
        items.setCount(3);

        //Set rowfields
        CTRowFields rowFields = definition.addNewRowFields();
        rowFields.addNewField().setX(0);
        rowFields.setCount(1);
        
        //Add rowItems
        CTRowItems rowItems = definition.addNewRowItems();
        rowItems.addNewI().addNewX();
        rowItems.addNewI().addNewX().setV(1);
        CTI rowItem = rowItems.addNewI();
        rowItem.setT(STItemType.GRAND);
        rowItem.addNewX();
        rowItems.setCount(3);
        
        //Set colItems
        CTColItems colItems = definition.addNewColItems();
        colItems.addNewI();
        colItems.setCount(1);
                
        //Set datafields, hard coded
        CTDataFields dataFields = definition.addNewDataFields();
        dataFields.setCount(1);
        CTDataField dataField = dataFields.addNewDataField();
        dataField.setName("Sum of #");
        //Index of the field to bee summarized
        dataField.setFld(1);
        
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
    }
}