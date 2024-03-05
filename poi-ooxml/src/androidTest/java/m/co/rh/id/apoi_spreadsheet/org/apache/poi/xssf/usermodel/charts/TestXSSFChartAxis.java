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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.charts;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;
import java.util.List;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.AxisPosition;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.AxisTickMark;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFChart;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFDrawing;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public final class TestXSSFChartAxis {

    private static final double EPSILON = 1E-7;
    private XSSFWorkbook wb;
    private XDDFChartAxis axis;

    @Before
    public void setup() {
        wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 10, 30);
        XSSFChart chart = drawing.createChart(anchor);
        axis = chart.createValueAxis(AxisPosition.BOTTOM);
        // no format set yet
        assertNull(axis.getNumberFormat());
    }

    @After
    public void teardown() throws IOException {
        wb.close();
        wb = null;
        axis = null;
    }

    @Test
    public void testLogBaseIllegalArgument() {
        IllegalArgumentException iae = null;
        try {
            axis.setLogBase(0.0);
        } catch (IllegalArgumentException e) {
            iae = e;
        }
        assertNotNull(iae);

        iae = null;
        try {
            axis.setLogBase(30000.0);
        } catch (IllegalArgumentException e) {
            iae = e;
        }
        assertNotNull(iae);
    }

    @Test
    public void testLogBaseLegalArgument() {
        axis.setLogBase(Math.E);
        assertTrue(Math.abs(axis.getLogBase() - Math.E) < EPSILON);
    }

    @Test
    public void testNumberFormat() {
        final String numberFormat = "General";
        axis.setNumberFormat(numberFormat);
        assertEquals(numberFormat, axis.getNumberFormat());
    }

    @Test
    public void testMaxAndMinAccessMethods() {
        final double newValue = 10.0;

        axis.setMinimum(newValue);
        assertTrue(Math.abs(axis.getMinimum() - newValue) < EPSILON);

        axis.setMaximum(newValue);
        assertTrue(Math.abs(axis.getMaximum() - newValue) < EPSILON);
    }

    @Test
    public void testVisibleAccessMethods() {
        axis.setVisible(true);
        assertTrue(axis.isVisible());

        axis.setVisible(false);
        assertFalse(axis.isVisible());
    }

    @Test
    public void testMajorTickMarkAccessMethods() {
        axis.setMajorTickMark(AxisTickMark.NONE);
        assertEquals(AxisTickMark.NONE, axis.getMajorTickMark());

        axis.setMajorTickMark(AxisTickMark.IN);
        assertEquals(AxisTickMark.IN, axis.getMajorTickMark());

        axis.setMajorTickMark(AxisTickMark.OUT);
        assertEquals(AxisTickMark.OUT, axis.getMajorTickMark());

        axis.setMajorTickMark(AxisTickMark.CROSS);
        assertEquals(AxisTickMark.CROSS, axis.getMajorTickMark());
    }

    @Test
    public void testMinorTickMarkAccessMethods() {
        axis.setMinorTickMark(AxisTickMark.NONE);
        assertEquals(AxisTickMark.NONE, axis.getMinorTickMark());

        axis.setMinorTickMark(AxisTickMark.IN);
        assertEquals(AxisTickMark.IN, axis.getMinorTickMark());

        axis.setMinorTickMark(AxisTickMark.OUT);
        assertEquals(AxisTickMark.OUT, axis.getMinorTickMark());

        axis.setMinorTickMark(AxisTickMark.CROSS);
        assertEquals(AxisTickMark.CROSS, axis.getMinorTickMark());
    }

    @Test
    public void testGetChartAxisBug57362() throws IOException {
        //Load existing excel with some chart on it having primary and secondary axis.
        try (final XSSFWorkbook workbook = XSSFTestDataSamples.openSampleWorkbook("57362.xlsx")) {
            final XSSFSheet sh = workbook.getSheetAt(0);
            final XSSFDrawing drawing = sh.createDrawingPatriarch();
            final XSSFChart chart = drawing.getCharts().get(0);

            final List<? extends XDDFChartAxis> axisList = chart.getAxes();

            assertEquals(4, axisList.size());
            assertNotNull(axisList.get(0));
            assertNotNull(axisList.get(1));
            assertNotNull(axisList.get(2));
            assertNotNull(axisList.get(3));
        }
    }
}
