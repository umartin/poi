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

import java.io.IOException;
import java.io.OutputStream;
import javax.xml.namespace.QName;
import org.apache.poi.POIXMLDocumentPart;
import static org.apache.poi.POIXMLDocumentPart.DEFAULT_XML_OPTIONS;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTI;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheRecords;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRecord;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STAxis;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;

/**
 *
 */
public class XSSFPivotTable extends POIXMLDocumentPart {
    
    protected final static short CREATED_VERSION = 3;
    protected final static short MIN_REFRESHABLE_VERSION = 3;
    protected final static short UPDATED_VERSION = 3;
    
    private CTPivotCache pivotCache;
    private CTPivotTableDefinition pivotTableDefinition;
    private XSSFPivotCacheDefinition pivotCacheDefinition;
    private XSSFPivotCacheRecords pivotCacheRecords;
    private XSSFSheet parentSheet;
    private int referenceStartRow;
    private int referenceEndRow;
    private int referenceStartColumn;
    private int referenceEndColumn;

    public XSSFPivotTable() {
        super();
        pivotTableDefinition = CTPivotTableDefinition.Factory.newInstance();
        pivotCache = CTPivotCache.Factory.newInstance();
        pivotCacheDefinition = new XSSFPivotCacheDefinition();
        pivotCacheRecords = new XSSFPivotCacheRecords();
        //setDefaultPivotTableDefinition();
    }

    public void setCache(CTPivotCache pivotCache) {
        this.pivotCache = pivotCache;
    }

    public CTPivotCache getCTPivotCache() {
        return pivotCache;
    }

    public XSSFSheet getParentSheet() {
        return parentSheet;
    }

    public void setParentSheet(XSSFSheet parentSheet) {
        this.parentSheet = parentSheet;
    }

    public CTPivotTableDefinition getCTPivotTableDefinition() {
        return pivotTableDefinition;
    }

    public void setCTPivotTableDefinition(CTPivotTableDefinition pivotTableDefinition) {
        this.pivotTableDefinition = pivotTableDefinition;
    }

    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
        return pivotCacheDefinition;
    }

    public void setPivotCacheDefinition(XSSFPivotCacheDefinition pivotCacheDefinition) {
        this.pivotCacheDefinition = pivotCacheDefinition;
    }

    public XSSFPivotCacheRecords getPivotCacheRecords() {
        return pivotCacheRecords;
    }

    public void setPivotCacheRecords(XSSFPivotCacheRecords pivotCacheRecords) {
        this.pivotCacheRecords = pivotCacheRecords;
    }
    
    /**
     * Set the area of where the values will be gathered.
     * Index starts at 0.
     * @param startColumn, the first column in the area.
     * @param startRow, the first row in the area.
     * @param endColumn, the last column in the area.
     * @param endRow, the last row in the area.
     */
    public void setReferences(int startColumn, int startRow, int endColumn, int endRow) {
        this.referenceStartColumn = startColumn;
        this.referenceStartRow = startRow;
        this.referenceEndColumn = endColumn;
        this.referenceEndRow = endRow;
    }
    
    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        //Sets the pivotTableDefinition tag
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTPivotTableDefinition.type.getName().
                getNamespaceURI(), "pivotTableDefinition"));
        pivotTableDefinition.save(out, xmlOptions);
        out.close();
    }
    
    /**
     * Set default values for the table definition.
     */
    public void setDefaultPivotTableDefinition() {
        //Not included
        //multipleFieldFilter (set when adding filter)
        //outlineData (set when grouping data)
        //outline -II-
        
        //Indentation increment for compact rows
        pivotTableDefinition.setIndent(0);
        //The pivot version which created the pivot cache set to default value
        pivotTableDefinition.setCreatedVersion(CREATED_VERSION);
        //Minimun version required to update the pivot cache
        pivotTableDefinition.setMinRefreshableVersion(MIN_REFRESHABLE_VERSION);
        //Version of the application which "updated the spreadsheet last"
        pivotTableDefinition.setUpdatedVersion(UPDATED_VERSION);
        //Titles shown at the top of each page when printed
        pivotTableDefinition.setItemPrintTitles(true);
        //Set autoformat properties      
        pivotTableDefinition.setUseAutoFormatting(true);
        pivotTableDefinition.setApplyNumberFormats(false);
        pivotTableDefinition.setApplyWidthHeightFormats(true);
        pivotTableDefinition.setApplyAlignmentFormats(false);
        pivotTableDefinition.setApplyPatternFormats(false);
        pivotTableDefinition.setApplyFontFormats(false);
        pivotTableDefinition.setApplyBorderFormats(false);
        pivotTableDefinition.setCacheId(pivotCache.getCacheId());
        pivotTableDefinition.setName("PivotTable"+pivotTableDefinition.getCacheId());
        pivotTableDefinition.setDataCaption("Values");
        
        //Set the default style for the pivot table
        CTPivotTableStyle style = pivotTableDefinition.addNewPivotTableStyleInfo();
        style.setName("PivotStyleLight16");
        style.setShowLastColumn(true);
        style.setShowColStripes(false);
        style.setShowRowStripes(false);
        style.setShowColHeaders(true);
        style.setShowRowHeaders(true);
    }
   
   /**
    * Set location of where the pivotTable will be placed.
    * @param ref, coordinates in worksheet, eg. A1:D4
    * @param firstDataCol, set which column is the first containing data (index of pivot table)
    * @param firstRowCol, set which row is the first containing data (index of pivot table)
    * @param firstHeaderRow, set which row is the first header row (index of pivot table)
    * @return the location of the table
    */
    public CTLocation setLocation(String ref, int firstDataCol, int firstDataRow, int firstHeaderRow) {
        CTLocation location;
        if(pivotTableDefinition.getLocation() == null) {
            location = pivotTableDefinition.addNewLocation();
            location.setFirstDataCol(firstDataCol);
            location.setFirstDataRow(firstDataRow);
            location.setFirstHeaderRow(firstHeaderRow);
        } else {
            location = pivotTableDefinition.getLocation();
        }
        location.setRef(ref);
        pivotTableDefinition.setLocation(location); 
        return location;
    }
    
    /**
     * Creates all pivotCacheRecords in the referenced area.
     */
    public void createCacheRecords() {
        CTPivotCacheRecords records =  pivotCacheRecords.getCtPivotCacheRecords();
        records.setCount(referenceEndRow-referenceStartRow);
        CTRecord record;
        Cell cell;
        Row row;
        //Goes through all cells, except the header, in the referenced area.
        for(int i = referenceStartRow+1; i <= referenceEndRow; i++) {
            row = parentSheet.getRow(i);
            record = records.addNewR();
            for(int j = referenceStartColumn; j <= referenceEndColumn; j++) {
                cell = row.getCell(j);
                //Creates a record based on the content of the cell.
                switch (cell.getCellType()) {
                    case (Cell.CELL_TYPE_BOOLEAN):
                        record.addNewB().setV(cell.getBooleanCellValue());
                        break;
                    case (Cell.CELL_TYPE_STRING):
                        record.addNewS().setV(cell.getStringCellValue());
                        break;
                    case (Cell.CELL_TYPE_NUMERIC):
                        record.addNewN().setV(cell.getNumericCellValue());
                        break;
                    case (Cell.CELL_TYPE_BLANK):
                        record.addNewM();
                        break;
                    /*case (Cell.CELL_TYPE_ERROR):
                        r.addNewE().setV(String.valueOf(c.getErrorCellValue()));
                        break;*/
                    case (Cell.CELL_TYPE_FORMULA):
                        record.addNewS().setV(cell.getCellFormula());
                        break;
                    default:
                        break;
                }
            }
        }   
    }
    
    public void addPivotFields() {
        CTPivotFields pivotFields = pivotTableDefinition.addNewPivotFields();     
        CTPivotField pivotField;

        for(long i = referenceStartColumn; i <= referenceEndColumn; i++) {
            pivotField = pivotFields.addNewPivotField();
            if((i-referenceStartColumn) == pivotTableDefinition.getLocation().getFirstDataCol()) {
                pivotField.setDataField(true);
            } else {
                CTItems items = pivotField.addNewItems();
                //Set dynamic?
                pivotField.setAxis(STAxis.AXIS_ROW);
                for(long j = referenceStartRow; j <= referenceEndRow; j++) {
                    items.addNewItem().setT(STItemType.DEFAULT);
                }   
                items.setCount(items.getItemList().size());
            }
        }        
        pivotFields.setCount(pivotFields.getPivotFieldList().size());       
    }
    
    public void addRowItems(AreaReference column) { 

        long startRow = column.getFirstCell().getRow();
        long endRow = column.getLastCell().getRow();     

        CTRowItems rowItems = pivotTableDefinition.addNewRowItems();

        for(long j = startRow; j <= endRow; j++) {
            if(j == endRow) {
                CTI rowItem = rowItems.addNewI();
                rowItem.setT(STItemType.GRAND);
                rowItem.addNewX(); 
            } else {
                 rowItems.addNewI().addNewX();
            }
        }
        rowItems.setCount(rowItems.getIList().size());
    }     
}