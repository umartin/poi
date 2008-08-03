/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.hssf.util;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.record.RecordInputStream;
import org.apache.poi.util.LittleEndian;

/**
 * Implementation of the cell range address lists,like is described
 * in OpenOffice.org's Excel Documentation: excelfileformat.pdf sec 2.5.14 -
 * 'Cell Range Address List'
 * 
 * In BIFF8 there is a common way to store absolute cell range address lists in
 * several records (not formulas). A cell range address list consists of a field
 * with the number of ranges and the list of the range addresses. Each cell
 * range address (called an ADDR structure) contains 4 16-bit-values.
 * </p>
 * 
 * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
 */
public final class CellRangeAddressList {

	/**
	 * List of <tt>CellRangeAddress</tt>es. Each structure represents a cell range
	 */
	private final List _list;

	public CellRangeAddressList() {
		_list = new ArrayList();
	}
	/**
	 * Convenience constructor for creating a <tt>CellRangeAddressList</tt> with a single 
	 * <tt>CellRangeAddress</tt>.  Other <tt>CellRangeAddress</tt>es may be added later.
	 */
	public CellRangeAddressList(int firstRow, int lastRow, int firstCol, int lastCol) {
		this();
		addCellRangeAddress(firstRow, firstCol, lastRow, lastCol);
	}

	/**
	 * @param in the RecordInputstream to read the record from
	 */
	public CellRangeAddressList(RecordInputStream in) {
		int nItems = in.readUShort();
		_list = new ArrayList(nItems);

		for (int k = 0; k < nItems; k++) {
			_list.add(new CellRangeAddress(in));
		}
	}

	/**
	 * Get the number of following ADDR structures. The number of this
	 * structures is automatically set when reading an Excel file and/or
	 * increased when you manually add a new ADDR structure . This is the reason
	 * there isn't a set method for this field .
	 * 
	 * @return number of ADDR structures
	 */
	public int getADDRStructureNumber() {
		return _list.size();
	}

	/**
	 * Add an ADDR structure .
	 * 
	 * @param firstRow - the upper left hand corner's row
	 * @param firstCol - the upper left hand corner's col
	 * @param lastRow - the lower right hand corner's row
	 * @param lastCol - the lower right hand corner's col
	 * @return the index of this ADDR structure
	 */
	public void addCellRangeAddress(int firstRow, int firstCol, int lastRow, int lastCol) {
		CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		_list.add(region);
	}

	/**
	 * @return <tt>CellRangeAddress</tt> at the given index
	 */
	public CellRangeAddress getCellRangeAddress(int index) {
		return (CellRangeAddress) _list.get(index);
	}

	public int serialize(int offset, byte[] data) {
		int pos = 2;

		int nItems = _list.size();
		LittleEndian.putUShort(data, offset, nItems);
		for (int k = 0; k < nItems; k++) {
			CellRangeAddress region = (CellRangeAddress) _list.get(k);
			pos += region.serialize(data, offset + pos);
		}
		return getSize();
	}

	public int getSize() {
		return 2 + CellRangeAddress.getEncodedSize(_list.size());
	}
}