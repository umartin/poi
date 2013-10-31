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
package org.apache.poi.xssf.usermodel;

import static junit.framework.TestCase.assertEquals;
import static junit.framework.TestCase.fail;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

public class TestXSSFPivotTable {
    
    XSSFPivotTable pivotTable;
    @Before
    public void setup(){
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet(); 
        
        Row row1 = sheet.createRow((short) 0);
        // Create a cell and put a value in it.
        Cell cell = row1.createCell((short) 0);
        cell.setCellValue("Names");
        Cell cell2 = row1.createCell((short) 1);
        cell2.setCellValue("#");
        Cell cell7 = row1.createCell((short) 2);
        cell7.setCellValue("Data");

        Row row2 = sheet.createRow((short) 1);
        Cell cell3 = row2.createCell((short) 0);
        cell3.setCellValue("Jan");
        Cell cell4 = row2.createCell((short) 1);
        cell4.setCellValue(10);
        Cell cell8 = row2.createCell((short) 2);
        cell8.setCellValue("Apa");

        Row row3 = sheet.createRow((short) 2);
        Cell cell5 = row3.createCell((short) 0);
        cell5.setCellValue("Ben");
        Cell cell6 = row3.createCell((short) 1);
        cell6.setCellValue(9);
        Cell cell9 = row3.createCell((short) 2);
        cell9.setCellValue("Bepa");  
        
        AreaReference source = new AreaReference("A1:B2");
        pivotTable = sheet.createPivotTable(source, new CellReference("H5"));
    }
    
    /*
     * Verify that when creating a row label it's  created on the correct row
     * and the count is increased by one.
     */
    @Test
    public void testAddRowLabelToPivotTable() {
        int columnIndex = 0;
        
        pivotTable.addRowLabel(columnIndex);
        CTPivotTableDefinition defintion = pivotTable.getCTPivotTableDefinition();
        
        assertEquals(defintion.getRowFields().getFieldArray(0).getX(), columnIndex);
        assertEquals(defintion.getRowFields().getCount(), 1);
    }
    /**
     * Verify that it's not possible to create a row label outside of the referenced area.
     */
    @Test
    public void testAddRowLabelOutOfRangeThrowsException() {
        int columnIndex = 5;
                
        try {
            pivotTable.addRowLabel(columnIndex);    
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        fail();
    }
    
     /**
     * Verify when creating a data column set to a data field, the data field with the corresponding
     * column index will be set to true.
     */
    @Test
    public void testAddDataColumn() {
        int columnIndex = 0;
        boolean isDataField = true;
        
        pivotTable.addDataColumn(columnIndex, isDataField);
        CTPivotFields pivotFields = pivotTable.getCTPivotTableDefinition().getPivotFields();
        assertEquals(pivotFields.getPivotFieldArray(columnIndex).getDataField(), isDataField);
    }
    
    /**
     * Verify that it's not possible to create a data column outside of the referenced area.
     */
    @Test
    public void testAddDataColumnOutOfRangeThrowsException() {         
        int columnIndex = 5;
        boolean isDataField = true;
        
        try {
            pivotTable.addDataColumn(columnIndex, isDataField);
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        fail(); 
    }
}
