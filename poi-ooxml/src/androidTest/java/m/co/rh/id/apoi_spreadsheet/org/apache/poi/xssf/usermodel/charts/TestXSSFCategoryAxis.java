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

import org.junit.Test;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.AxisPosition;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFChart;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFDrawing;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public final class TestXSSFCategoryAxis {
    @Test
    public void testAccessMethods() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet();
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            XSSFChart chart = drawing.createChart(anchor);
            XDDFCategoryAxis axis = chart.createCategoryAxis(AxisPosition.BOTTOM);

            axis.setCrosses(AxisCrosses.AUTO_ZERO);
            assertEquals(AxisCrosses.AUTO_ZERO, axis.getCrosses());

            assertEquals(1, chart.getAxes().size());
        }
    }
}
