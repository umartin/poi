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
import java.util.List;
import javax.xml.namespace.QName;
import org.apache.poi.POIXMLDocumentPart;
import static org.apache.poi.POIXMLDocumentPart.DEFAULT_XML_OPTIONS;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheRecords;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRecord;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STAxis;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSourceType;

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

    public XSSFPivotTable() {
        super();
        pivotTableDefinition = CTPivotTableDefinition.Factory.newInstance();
        pivotCache = CTPivotCache.Factory.newInstance();
        pivotCacheDefinition = new XSSFPivotCacheDefinition();
        pivotCacheRecords = new XSSFPivotCacheRecords();
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
    protected void setDefaultPivotTableDefinition() {
        //Not more than one until more created
        pivotTableDefinition.setMultipleFieldFilters(false);
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
     * Creates all pivotCacheRecords in the referenced area.
     */
    protected void createCacheRecords() {
        CTPivotCacheRecords records =  pivotCacheRecords.getCtPivotCacheRecords();
        String source = pivotCacheDefinition.getCTPivotCacheDefinition().
                getCacheSource().getWorksheetSource().getRef();
        AreaReference sourceArea = new AreaReference(source);

        CTRecord record;
        Cell cell;
        Row row;
        //Goes through all cells, except the header, in the referenced area.
        for(int i = sourceArea.getFirstCell().getRow()+1; i <= sourceArea.getLastCell().getRow(); i++) {
            row = parentSheet.getRow(i);
            record = records.addNewR();
            for(int j = sourceArea.getFirstCell().getCol(); j <= sourceArea.getLastCell().getCol(); j++) {
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
                    case (Cell.CELL_TYPE_ERROR):
                        record.addNewE().setV(String.valueOf(cell.getErrorCellValue()));
                        break;
                    case (Cell.CELL_TYPE_FORMULA):
                        record.addNewS().setV(cell.getCellFormula());
                        break;
                    default:
                        break;
                }
            }
        }                  
        records.setCount(records.getRList().size());
    }
    
    /**
     * Add a row label using data from the given column.
     * @param columnIndex, the index of the column to be used as row label.
     */
    public void addRowLabel(int columnIndex) {
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();            

        AreaReference pivotArea = new AreaReference(pivotTableDefinition.getLocation().getRef());
        int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
    
        List<CTPivotField> pivotFieldList = pivotTableDefinition.getPivotFields().getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        CTItems items = pivotField.addNewItems();

        pivotField.setAxis(STAxis.AXIS_ROW);
        for(int i = 0; i < lastRowIndex; i++) {
            items.addNewItem().setT(STItemType.DEFAULT);
        }
        items.setCount(items.getItemList().size());
        pivotFieldList.set(columnIndex, pivotField);
        
        pivotFields.setPivotFieldArray(pivotFieldList.toArray(new CTPivotField[pivotFieldList.size()]));
        
        CTRowFields rowFields;
        if(pivotTableDefinition.getRowFields() != null) {
            rowFields = pivotTableDefinition.getRowFields();
        } else {
            rowFields = pivotTableDefinition.addNewRowFields();
        }
        
        rowFields.addNewField().setX(columnIndex);
        rowFields.setCount(rowFields.getFieldList().size());
    }
    
    /**
     * Add a column label using data from the given column and specified function
     * @param columnIndex, the index of the column to be used as column label.
     * @param function, the function to be used on the data
     * The following functions exists:
     * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
     */
    public void addColumnLabel(STDataConsolidateFunction.Enum function, int columnIndex) {
        addDataColumn(columnIndex, true);       
        addDataField(function, columnIndex);
        if (pivotTableDefinition.getDataFields().getCount() > 1) {
            CTColFields colFields;
            if(pivotTableDefinition.getColFields() != null) {
                colFields = pivotTableDefinition.getColFields();    
            } else {
                colFields = pivotTableDefinition.addNewColFields();
            }     
            colFields.addNewField().setX(-2);
            colFields.setCount(colFields.getFieldList().size());
        }
    }
    
    /**
     * Add data field with data from the given column and specified function.
     * @param function, the function to be used on the data
     * @param index, the index of the column to be used as column label.
     * The following functions exists:
     * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
     */
    private void addDataField(STDataConsolidateFunction.Enum function, int index) {
        CTDataFields dataFields;
        if(pivotTableDefinition.getDataFields() != null) {
            dataFields = pivotTableDefinition.getDataFields();
        } else {
            dataFields = pivotTableDefinition.addNewDataFields();
        }
        CTDataField dataField = dataFields.addNewDataField();
        dataField.setSubtotal(function);
        dataField.setFld(index);          
        dataFields.setCount(dataFields.getDataFieldList().size());
    }
    
    /**
     * All columns in the referenced area must be added to the pivot table, either as
     * a column/row label or as a data column. Not all data columns must be displayed in the
     * pivot table.
     * 
     * Add column containing data from the referenced area.
     * @param columnIndex, the index of the column containing the data
     * @param isDataField, true if the data should be displayed in the pivot table.
     */
    public void addDataColumn(int columnIndex, boolean isDataField) {
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();            
        List<CTPivotField> pivotFieldList = pivotFields.getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        
        pivotField.setDataField(isDataField);         
        pivotFieldList.set(columnIndex, pivotField);
        pivotFields.setPivotFieldArray(pivotFieldList.toArray(new CTPivotField[pivotFieldList.size()]));
    }
    
    /**
     * Add filter for the column with the corresponding index 
     * @param index, of the column to filter on
     */
    public void addReportFilter(int index) {
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();
        AreaReference pivotArea = new AreaReference(pivotTableDefinition.getLocation().getRef());
        int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
    
        List<CTPivotField> pivotFieldList = pivotTableDefinition.getPivotFields().getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        CTItems items = pivotField.addNewItems();

        pivotField.setAxis(STAxis.AXIS_PAGE);
        for(int i = 0; i < lastRowIndex; i++) {
            items.addNewItem().setT(STItemType.DEFAULT);
        }
        items.setCount(items.getItemList().size());
        pivotFieldList.set(index, pivotField);
        
        CTPageFields pageFields;
        if (pivotTableDefinition.getPageFields()!= null) {
            pageFields = pivotTableDefinition.getPageFields();  
            //Another filter has already been created
            pivotTableDefinition.setMultipleFieldFilters(true);
        } else {
            pageFields = pivotTableDefinition.addNewPageFields();
        }
        
        CTPageField pageField = pageFields.addNewPageField();
        pageField.setHier(-1);
        pageField.setFld(index);
    }
    
    /**
     * Creates cacheSource and workSheetSource for pivot table and sets the source reference as well assets the location of the pivot table
     * @param source Source for data for pivot table
     * @param position Position for pivot table in sheet
     * @param sourceSheet Sheet where the source will be collected from
     */
    protected void createSourceReferences(AreaReference source, CellReference position, XSSFSheet sourceSheet){
        //Get cell one to the right and one down from position, add both to AreaReference and set pivot table location.
        AreaReference destination = new AreaReference(position, new CellReference(position.getRow()+1, position.getCol()+1));
        
        CTLocation location;
        if(pivotTableDefinition.getLocation() == null) {
            location = pivotTableDefinition.addNewLocation();
            location.setFirstDataCol(1);
            location.setFirstDataRow(1);
            location.setFirstHeaderRow(1);
        } else {
            location = pivotTableDefinition.getLocation();
        }
        location.setRef(destination.formatAsString());
        pivotTableDefinition.setLocation(location); 

        //Set source for the pivot table
        CTPivotCacheDefinition cacheDef = getPivotCacheDefinition().getCTPivotCacheDefinition();
        CTCacheSource cacheSource = cacheDef.addNewCacheSource();
        cacheSource.setType(STSourceType.WORKSHEET);
        CTWorksheetSource worksheetSource = cacheSource.addNewWorksheetSource();
        worksheetSource.setSheet(sourceSheet.getSheetName());
        worksheetSource.setRef(source.formatAsString());
    }
    
    protected void createDefaultDataColumns() {
        CTPivotFields pivotFields;
        if (pivotTableDefinition.getPivotFields() != null) {
            pivotFields = pivotTableDefinition.getPivotFields();            
        } else {
            pivotFields = pivotTableDefinition.addNewPivotFields();
        }
        String source = pivotCacheDefinition.getCTPivotCacheDefinition().
                getCacheSource().getWorksheetSource().getRef();
        AreaReference sourceArea = new AreaReference(source);
        int firstColumn = sourceArea.getFirstCell().getCol();
        int lastColumn = sourceArea.getLastCell().getCol();
        for(int i = 0; i<=lastColumn-firstColumn; i++) {
            pivotFields.addNewPivotField().setDataField(false);
        }
        pivotFields.setCount(pivotFields.getPivotFieldList().size());
    }
}