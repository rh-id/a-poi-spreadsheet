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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DifferentialStyleProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.PatternFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Table;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.TableStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.TableStyleInfo;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.TableStyleType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;

/**
 * Test built-in table styles
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestTableStyles {

    /**
     * Test that a built-in style is initialized properly
     */
    @Test
    public void testBuiltinStyleInit() {
        TableStyle style = XSSFBuiltinTableStyle.TableStyleMedium2.getStyle();
        assertNotNull("no style found for Medium2", style);
        assertNull("Should not have style info for blankRow", style.getStyle(TableStyleType.blankRow));
        DifferentialStyleProvider headerRow = style.getStyle(TableStyleType.headerRow);
        assertNotNull("no header row style", headerRow);
        FontFormatting font = headerRow.getFontFormatting();
        assertNotNull("No header row font formatting", font);
        assertTrue("header row not bold", font.isBold());
        PatternFormatting fill = headerRow.getPatternFormatting();
        assertNotNull("No header fill", fill);
        assertEquals("wrong header fill",
                4, ((XSSFColor) fill.getFillBackgroundColorColor()).getTheme());
    }

    @Test
    public void testCustomStyle() throws Exception {
        try (XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("tableStyle.xlsx")) {
            Table table = wb.getTable("Table1");
            assertNotNull("missing table", table);

            TableStyleInfo style = table.getStyle();
            assertNotNull("Missing table style info", style);
            assertNotNull("Missing table style", style.getStyle());
            assertEquals("Wrong name", "TestTableStyle", style.getName());
            assertEquals("Wrong name", "TestTableStyle", style.getStyle().getName());

            DifferentialStyleProvider firstColumn = style.getStyle().getStyle(TableStyleType.firstColumn);
            assertNotNull("no first column style", firstColumn);
            FontFormatting font = firstColumn.getFontFormatting();
            assertNotNull("no first col font", font);
            assertTrue("wrong first col bold", font.isBold());
        }
    }
}
