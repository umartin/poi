package org.apache.poi.xssf.usermodel;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.junit.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;

/**
 *
 * @author Martin Andersson
 */
public class XSSFPivotDataTest {

	private static final int DATA_ROWS = 200;
	private static final String[] NAMES = new String[] {"Kalle", "Ove", "Pelle"};

	@Test
	public void testAddDataAfterPivotIsCreated() throws FileNotFoundException, IOException {

		Workbook wb = new XSSFWorkbook();
		XSSFSheet pivotSheet = (XSSFSheet) wb.createSheet("pivot");
		XSSFSheet dataSheet = (XSSFSheet) wb.createSheet("data");

		// Only create data headers.
		dataSheet.createRow(0).createCell(0).setCellValue("Name");
		dataSheet.getRow(0).createCell(1).setCellValue("Age");

		// Add pivot table.
		AreaReference pivotDataArea = new AreaReference(new CellReference("A1"), new CellReference(DATA_ROWS , 1));

		XSSFPivotTable pivotTable = pivotSheet.createPivotTable(pivotDataArea, new CellReference("A3"), dataSheet);

		pivotTable.addRowLabel(0);
		pivotTable.addColumnLabel(STDataConsolidateFunction.AVERAGE, 1);

		// Add data after the pivot is created.
		addData(dataSheet);
	}

	private void addData(Sheet sheet) {
		for (int i = 0; i < DATA_ROWS; i++) {
			Row row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(NAMES[i % NAMES.length]);
			row.createCell(1).setCellValue(i);
		}
	}
}
