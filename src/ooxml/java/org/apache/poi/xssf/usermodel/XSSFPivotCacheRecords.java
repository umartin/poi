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
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheRecords;

public class XSSFPivotCacheRecords extends POIXMLDocumentPart {
    
    private CTPivotCacheRecords ctPivotCacheRecords;
    
    public XSSFPivotCacheRecords(){
        super();
        ctPivotCacheRecords = CTPivotCacheRecords.Factory.newInstance();
    }
    
    /**
     * Creates an XSSFPivotCacheRecords representing the given package part and relationship.
     *
     * @param part - The package part that holds xml data representing this pivot cache records.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    protected XSSFPivotCacheRecords(PackagePart part, PackageRelationship rel) {
        super(part, rel);
        ctPivotCacheRecords = CTPivotCacheRecords.Factory.newInstance();
    }

    public CTPivotCacheRecords getCtPivotCacheRecords() {
        return ctPivotCacheRecords;
    }
    
    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        //Sets the pivotCacheRecords tag
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTPivotCacheRecords.type.getName().
                getNamespaceURI(), "pivotCacheRecords"));
        ctPivotCacheRecords.save(out, xmlOptions);
        out.close();
    }
    
}
