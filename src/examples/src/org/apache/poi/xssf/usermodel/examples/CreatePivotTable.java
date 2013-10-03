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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/**
 *
 */
public class CreatePivotTable {
    
    public static void main(String[] args) throws FileNotFoundException, IOException {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();
        XSSFPivotTable pivotTable = sheet.createPivotTable();
        
        CTPivotTableDefinition pivotTableDefinition = pivotTable.getCTPivotTableDefinition();
        pivotTableDefinition.setMultipleFieldFilters(false);
        pivotTableDefinition.setOutlineData(true);
        pivotTableDefinition.setOutline(true);
        pivotTableDefinition.setIndent(0);
        pivotTableDefinition.setCreatedVersion(new Short("4"));
        pivotTableDefinition.setItemPrintTitles(true);
        pivotTableDefinition.setUseAutoFormatting(true);
        pivotTableDefinition.setMinRefreshableVersion(new Short("3"));
        pivotTableDefinition.setUpdatedVersion(new Short("4"));
        pivotTableDefinition.setDataCaption("Values");
        pivotTableDefinition.setApplyWidthHeightFormats(true);
        pivotTableDefinition.setApplyAlignmentFormats(false);
        pivotTableDefinition.setApplyPatternFormats(false);
        pivotTableDefinition.setApplyFontFormats(false);
        pivotTableDefinition.setApplyBorderFormats(false);
        pivotTableDefinition.setApplyNumberFormats(false);
        pivotTableDefinition.setDataOnRows(true);
        pivotTableDefinition.setCacheId(pivotTable.getCTPivotCache().getCacheId());
        pivotTableDefinition.setName("PivotTable1");

        FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.zip");
        wb.write(fileOut);
        fileOut.close(); 
    }

}
