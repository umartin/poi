package org.apache.poi.xssf.usermodel;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;

/**
 *
 */
public class XSSFPivotTable extends POIXMLDocumentPart{
    private CTPivotCache pivotCache;
    private CTPivotTableDefinition pivotTableDefinition;
    private String workbookRelationId;
    public XSSFPivotTable() {
        
    }

    public void setCache(CTPivotCache pivotCache) {
        this.pivotCache = pivotCache;
    }

    public CTPivotCache getPivotCache() {
        return pivotCache;
    }

    public CTPivotTableDefinition getPivotTableDefinition() {
        return pivotTableDefinition;
    }

    public void setPivotTableDefinition(CTPivotTableDefinition pivotTableDefinition) {
        this.pivotTableDefinition = pivotTableDefinition;
    }

    public String getWorkbookRelationId() {
        return workbookRelationId;
    }

    public void setWorkbookRelationId(String workbookRelationId) {
        this.workbookRelationId = workbookRelationId;
    }
    
}