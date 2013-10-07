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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSourceType;

public class XSSFPivotCacheDefinition extends POIXMLDocumentPart{
    
    private CTPivotCacheDefinition ctPivotCacheDefinition;
    
    public XSSFPivotCacheDefinition(){
        super();
        ctPivotCacheDefinition = CTPivotCacheDefinition.Factory.newInstance();
        createDefaultValues();
    }
    
    public CTPivotCacheDefinition getCTPivotCacheDefinition(){
        return ctPivotCacheDefinition;
    }
    
    private void createDefaultValues(){
        ctPivotCacheDefinition.setCreatedVersion(XSSFPivotTable.CREATED_VERSION);
        ctPivotCacheDefinition.setMinRefreshableVersion(XSSFPivotTable.MIN_REFRESHABLE_VERSION);
        ctPivotCacheDefinition.setRefreshedVersion(XSSFPivotTable.UPDATED_VERSION);
    }
    
    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        //Sets the pivotCacheDefinition tag
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTPivotCacheDefinition.type.getName().
                getNamespaceURI(), "pivotCacheDefinition"));
        ctPivotCacheDefinition.save(out, xmlOptions);
        out.close();
    }
}
