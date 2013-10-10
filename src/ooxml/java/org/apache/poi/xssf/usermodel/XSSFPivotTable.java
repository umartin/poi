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
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

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
    private XSSFPivotCacheRecords pivotCacheRecord;
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
        pivotCacheRecord = new XSSFPivotCacheRecords();
        setDefaultPivotTableDefinition();
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
        return pivotCacheRecord;
    }

    public void setPivotCacheRecords(XSSFPivotCacheRecords pivotCacheRecord) {
        this.pivotCacheRecord = pivotCacheRecord;
    }
    
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
    
    private void setDefaultPivotTableDefinition() {
        //Not included
        //multipleFieldFilter (set when adding filter)
        //outlineData (set when when grouping data)
        //outline -II-
        //dataCaption, the anme of the values area header
        
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
    }
   
    public CTLocation setLocation(String ref) {
        CTLocation location;
        if(pivotTableDefinition.getLocation() == null) {
            location = pivotTableDefinition.addNewLocation();
            //Provide some default 
            location.setFirstDataCol(1);
            location.setFirstDataRow(1);
            location.setFirstHeaderRow(1);
        } else {
            location = pivotTableDefinition.getLocation();
        }
        location.setRef(ref);
        pivotTableDefinition.setLocation(location); 
        return location;
    }
    
}