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

package org.apache.poi.xssf.usermodel.charts;

import junit.framework.TestCase;

import org.apache.poi.ss.usermodel.charts.TitleType;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Martin Andersson
 */
public class TestAbstractXSSFChartSerie extends TestCase {

	public void testTitleAccessorMethods() {
		AbstractXSSFChartSerie serie = new AbstractXSSFChartSerie() {};

		assertFalse(serie.isTitleSet());

		serie.setTitle("title");
		assertTrue(serie.isTitleSet());
		assertNotNull(serie.getCTSerTx());
		assertEquals(TitleType.STRING, serie.getTitleType());
		assertEquals("title", serie.getTitleString());

		CellReference cellRef = new CellReference("Sheet1!A1");
		serie.setTitle(cellRef);
		assertTrue(serie.isTitleSet());
		assertNotNull(serie.getCTSerTx());
		assertEquals(TitleType.CELL_REFERENCE, serie.getTitleType());
		assertEquals(cellRef, serie.getTitleCellReference());
	}
}
