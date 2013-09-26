package org.apache.poi.xssf.usermodel;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;

/**
 *
 */
public class XSSFPivotTable extends POIXMLDocumentPart{
    
    public XSSFPivotTable() {
    }
    
    public XSSFPivotTable(PackagePart part, PackageRelationship rel) {
        super(part, rel);
    }
    
}