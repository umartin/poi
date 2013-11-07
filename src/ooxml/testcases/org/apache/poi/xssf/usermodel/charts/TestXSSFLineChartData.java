package org.apache.poi.xssf.usermodel.charts;

import junit.framework.TestCase;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LineChartSerie;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetBuilder;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Martin Andersson
 */
public class TestXSSFLineChartData extends TestCase {

    private static final Object[][] plotData = {
            {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"},
            {  1,    2,   3,    4,    5,   6,    7,   8,    9,  10}
    };

    public void testOneSeriePlot() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = new SheetBuilder(wb, plotData).build();
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 10, 30);
        Chart chart = drawing.createChart(anchor);

        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ChartAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);

        LineChartData lineChartData =
                chart.getChartDataFactory().createLineChartData();

        ChartDataSource<String> xs = DataSources.fromStringCellRange(sheet, CellRangeAddress.valueOf("A1:J1"));
        ChartDataSource<Number> ys = DataSources.fromNumericCellRange(sheet, CellRangeAddress.valueOf("A2:J2"));
        LineChartSerie serie = lineChartData.addSerie(xs, ys);

        assertNotNull(serie);
        assertEquals(1, lineChartData.getSeries().size());
        assertTrue(lineChartData.getSeries().contains(serie));

        chart.plot(lineChartData, bottomAxis, leftAxis);
    }
}
