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
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import android.util.Log;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;

/**
 * Test asserts the POI produces &lt;cols&gt; element that could be read and properly interpreted by the MS Excel.
 * For specification of the "cols" element see the chapter 3.3.1.16 of the "Office Open XML Part 4 - Markup Language Reference.pdf".
 * The specification can be downloaded at https://www.ecma-international.org/publications/files/ECMA-ST/Office%20Open%20XML%201st%20edition%20Part%204%20(PDF).zip.
 *
 * <p><em>
 * The test saves xlsx file on a disk if the system property is set:
 * -Dpoi.test.xssf.output.dir=${workspace_loc}/poi/build/xssf-output
 * </em>
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFColGrouping {

    private static final String TAG = "TestXSSFColGrouping";


    /**
     * Tests that POI doesn't produce "col" elements without "width" attribute.
     * POI-52186
     */
    @Test
    public void testNoColsWithoutWidthWhenGrouping() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            sheet.setColumnWidth(4, 5000);
            sheet.setColumnWidth(5, 5000);

            sheet.groupColumn((short) 4, (short) 7);
            sheet.groupColumn((short) 9, (short) 12);

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testNoColsWithoutWidthWhenGrouping")) {
                sheet = wb2.getSheet("test");

                CTCols cols = sheet.getCTWorksheet().getColsArray(0);
                Log.d(TAG, String.format("test52186/cols:%s", cols));
                for (CTCol col : cols.getColArray()) {
                    assertTrue("Col width attribute is unset: " + col,
                            col.isSetWidth());
                }

            }
        }
    }

    /**
     * Tests that POI doesn't produce "col" elements without "width" attribute.
     * POI-52186
     */
    @Test
    public void testNoColsWithoutWidthWhenGroupingAndCollapsing() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            sheet.setColumnWidth(4, 5000);
            sheet.setColumnWidth(5, 5000);

            sheet.groupColumn((short) 4, (short) 5);

            sheet.setColumnGroupCollapsed(4, true);

            CTCols cols = sheet.getCTWorksheet().getColsArray(0);
            Log.d(TAG, String.format("test52186_2/cols:%s", cols));

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testNoColsWithoutWidthWhenGroupingAndCollapsing")) {
                sheet = wb2.getSheet("test");

                for (int i = 4; i <= 5; i++) {
                    assertEquals("Unexpected width of column " + i,
                            5000, sheet.getColumnWidth(i));
                }
                cols = sheet.getCTWorksheet().getColsArray(0);
                for (CTCol col : cols.getColArray()) {
                    assertTrue("Col width attribute is unset: " + col,
                            col.isSetWidth());
                }
            }
        }
    }

    /**
     * Test the cols element is correct in case of NumericRanges.OVERLAPS_2_WRAPS
     */
    @Test
    public void testMergingOverlappingCols_OVERLAPS_2_WRAPS() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            CTCols cols = sheet.getCTWorksheet().getColsArray(0);
            CTCol col = cols.addNewCol();
            col.setMin(1 + 1);
            col.setMax(4 + 1);
            col.setWidth(20);
            col.setCustomWidth(true);

            sheet.groupColumn((short) 2, (short) 3);

            sheet.getCTWorksheet().getColsArray(0);
            Log.d(TAG, String.format("testMergingOverlappingCols_OVERLAPS_2_WRAPS/cols:%s", cols));

            assertEquals(0, cols.getColArray(0).getOutlineLevel());
            assertEquals(2, cols.getColArray(0).getMin()); // 1 based
            assertEquals(2, cols.getColArray(0).getMax()); // 1 based
            assertTrue(cols.getColArray(0).getCustomWidth());

            assertEquals(1, cols.getColArray(1).getOutlineLevel());
            assertEquals(3, cols.getColArray(1).getMin()); // 1 based
            assertEquals(4, cols.getColArray(1).getMax()); // 1 based
            assertTrue(cols.getColArray(1).getCustomWidth());

            assertEquals(0, cols.getColArray(2).getOutlineLevel());
            assertEquals(5, cols.getColArray(2).getMin()); // 1 based
            assertEquals(5, cols.getColArray(2).getMax()); // 1 based
            assertTrue(cols.getColArray(2).getCustomWidth());

            assertEquals(3, cols.sizeOfColArray());

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testMergingOverlappingCols_OVERLAPS_2_WRAPS")) {
                sheet = wb2.getSheet("test");

                for (int i = 1; i <= 4; i++) {
                    assertEquals("Unexpected width of column " + i,
                            20 * 256, sheet.getColumnWidth(i));
                }

            }
        }
    }

    /**
     * Test the cols element is correct in case of NumericRanges.OVERLAPS_1_WRAPS
     */
    @Test
    public void testMergingOverlappingCols_OVERLAPS_1_WRAPS() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            CTCols cols = sheet.getCTWorksheet().getColsArray(0);
            CTCol col = cols.addNewCol();
            col.setMin(2 + 1);
            col.setMax(4 + 1);
            col.setWidth(20);
            col.setCustomWidth(true);

            sheet.groupColumn((short) 1, (short) 5);

            cols = sheet.getCTWorksheet().getColsArray(0);
            Log.d(TAG, String.format("testMergingOverlappingCols_OVERLAPS_1_WRAPS/cols:%s", cols));

            assertEquals(1, cols.getColArray(0).getOutlineLevel());
            assertEquals(2, cols.getColArray(0).getMin()); // 1 based
            assertEquals(2, cols.getColArray(0).getMax()); // 1 based
            assertFalse(cols.getColArray(0).getCustomWidth());

            assertEquals(1, cols.getColArray(1).getOutlineLevel());
            assertEquals(3, cols.getColArray(1).getMin()); // 1 based
            assertEquals(5, cols.getColArray(1).getMax()); // 1 based
            assertTrue(cols.getColArray(1).getCustomWidth());

            assertEquals(1, cols.getColArray(2).getOutlineLevel());
            assertEquals(6, cols.getColArray(2).getMin()); // 1 based
            assertEquals(6, cols.getColArray(2).getMax()); // 1 based
            assertFalse(cols.getColArray(2).getCustomWidth());

            assertEquals(3, cols.sizeOfColArray());

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testMergingOverlappingCols_OVERLAPS_1_WRAPS")) {
                sheet = wb2.getSheet("test");

                for (int i = 2; i <= 4; i++) {
                    assertEquals("Unexpected width of column " + i,
                            20 * 256, sheet.getColumnWidth(i));
                }

            }
        }
    }

    /**
     * Test the cols element is correct in case of NumericRanges.OVERLAPS_1_MINOR
     */
    @Test
    public void testMergingOverlappingCols_OVERLAPS_1_MINOR() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            CTCols cols = sheet.getCTWorksheet().getColsArray(0);
            CTCol col = cols.addNewCol();
            col.setMin(2 + 1);
            col.setMax(4 + 1);
            col.setWidth(20);
            col.setCustomWidth(true);

            sheet.groupColumn((short) 3, (short) 5);

            cols = sheet.getCTWorksheet().getColsArray(0);
            Log.d(TAG, String.format("testMergingOverlappingCols_OVERLAPS_1_MINOR/cols:%s", cols));

            assertEquals(0, cols.getColArray(0).getOutlineLevel());
            assertEquals(3, cols.getColArray(0).getMin()); // 1 based
            assertEquals(3, cols.getColArray(0).getMax()); // 1 based
            assertTrue(cols.getColArray(0).getCustomWidth());

            assertEquals(1, cols.getColArray(1).getOutlineLevel());
            assertEquals(4, cols.getColArray(1).getMin()); // 1 based
            assertEquals(5, cols.getColArray(1).getMax()); // 1 based
            assertTrue(cols.getColArray(1).getCustomWidth());

            assertEquals(1, cols.getColArray(2).getOutlineLevel());
            assertEquals(6, cols.getColArray(2).getMin()); // 1 based
            assertEquals(6, cols.getColArray(2).getMax()); // 1 based
            assertFalse(cols.getColArray(2).getCustomWidth());

            assertEquals(3, cols.sizeOfColArray());

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testMergingOverlappingCols_OVERLAPS_1_MINOR")) {
                sheet = wb2.getSheet("test");

                for (int i = 2; i <= 4; i++) {
                    assertEquals("Unexpected width of column " + i,
                            20 * 256L, sheet.getColumnWidth(i));
                }
                assertEquals("Unexpected width of column " + 5,
                        sheet.getDefaultColumnWidth() * 256L, sheet.getColumnWidth(5));

            }
        }
    }

    /**
     * Test the cols element is correct in case of NumericRanges.OVERLAPS_2_MINOR
     */
    @Test
    public void testMergingOverlappingCols_OVERLAPS_2_MINOR() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            XSSFSheet sheet = wb1.createSheet("test");

            CTCols cols = sheet.getCTWorksheet().getColsArray(0);
            CTCol col = cols.addNewCol();
            col.setMin(2 + 1);
            col.setMax(4 + 1);
            col.setWidth(20);
            col.setCustomWidth(true);

            sheet.groupColumn((short) 1, (short) 3);

            cols = sheet.getCTWorksheet().getColsArray(0);
            Log.d(TAG, String.format("testMergingOverlappingCols_OVERLAPS_2_MINOR/cols:%s", cols));

            assertEquals(1, cols.getColArray(0).getOutlineLevel());
            assertEquals(2, cols.getColArray(0).getMin()); // 1 based
            assertEquals(2, cols.getColArray(0).getMax()); // 1 based
            assertFalse(cols.getColArray(0).getCustomWidth());

            assertEquals(1, cols.getColArray(1).getOutlineLevel());
            assertEquals(3, cols.getColArray(1).getMin()); // 1 based
            assertEquals(4, cols.getColArray(1).getMax()); // 1 based
            assertTrue(cols.getColArray(1).getCustomWidth());

            assertEquals(0, cols.getColArray(2).getOutlineLevel());
            assertEquals(5, cols.getColArray(2).getMin()); // 1 based
            assertEquals(5, cols.getColArray(2).getMax()); // 1 based
            assertTrue(cols.getColArray(2).getCustomWidth());

            assertEquals(3, cols.sizeOfColArray());

            try (XSSFWorkbook wb2 = XSSFTestDataSamples.writeOutAndReadBack(wb1, "testMergingOverlappingCols_OVERLAPS_2_MINOR")) {
                sheet = wb2.getSheet("test");

                for (int i = 2; i <= 4; i++) {
                    assertEquals("Unexpected width of column " + i,
                            20 * 256L, sheet.getColumnWidth(i));
                }
                assertEquals("Unexpected width of column " + 1,
                        sheet.getDefaultColumnWidth() * 256L, sheet.getColumnWidth(1));

            }
        }
    }

}
