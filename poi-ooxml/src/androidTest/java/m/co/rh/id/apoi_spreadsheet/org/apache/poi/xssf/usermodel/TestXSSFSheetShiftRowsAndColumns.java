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
import static org.junit.Assert.assertTrue;

import androidx.test.platform.app.InstrumentationRegistry;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Locale;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFSheetShiftRowsAndColumns {
    private File resultDir;

    private static final int numRows = 4;
    private static final int numCols = 4;

    private static final int INSERT_ROW = 1;
    private static final int INSERT_COLUMN = 1;
    private static final int FIRST_MERGE_ROW = INSERT_ROW + 1;
    private static final int LAST_MERGE_ROW = numRows - 1;
    private static final int FIRST_MERGE_COL = INSERT_COLUMN + 1;
    private static final int LAST_MERGE_COL = numCols - 1;

    private XSSFWorkbook workbook = null;
    private XSSFSheet sheet = null;
    private String fileName = null;

    /**
     * This creates a workbook with one worksheet.  It then puts data in rows 0 to numRows-1 and colulmns
     * 0 to numCols-1.
     */
    @Before
    public void setup() throws IOException {
        File cache = InstrumentationRegistry.getInstrumentation().getTargetContext().getCacheDir();
        resultDir = new File(cache, "build/custom-reports-test");
        resultDir.mkdirs();
        assertTrue("Failed to create directory " + resultDir,
                resultDir.exists() || resultDir.mkdirs());
        final String procName = "TestXSSFSheetShiftRowsAndColumns.setup";
        fileName = procName + ".xlsx";

        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet();

        for (int nRow = 0; nRow < numRows; ++nRow) {
            final XSSFRow row = sheet.createRow(nRow);
            for (int nCol = 0; nCol < numCols; ++nCol) {
                final XSSFCell cell = row.createCell(nCol);
                cell.setCellType(CellType.STRING);
                cell.setCellValue(String.format(Locale.US, "Row %d col %d", nRow, nCol));
            }
        }
        /*
         * Add a merge area
         */
        final CellRangeAddress range = new CellRangeAddress(FIRST_MERGE_ROW, LAST_MERGE_ROW, FIRST_MERGE_COL, LAST_MERGE_COL);
        sheet.addMergedRegion(range);
        writeFile(procName);
    }


    /**
     * This method writes the workbook to resultDir/fileName.
     */
    @After
    public void cleanup() throws IOException {
        final String procName = "TestXSSFSheetRemoveTable.cleanup";
        if (workbook == null) {
            return;
        }

        if (fileName == null) {
            return;
        }

        writeFile(procName);

        workbook.close();
    }

    private void writeFile(String procName) throws IOException {
        final File file = new File(resultDir, fileName);
        try (OutputStream fileOut = new FileOutputStream(file)) {
            workbook.write(fileOut);
        }
    }

    /**
     * Apply no shift.  The purpose of this is to test {@code testCellAddresses} and {@code testMergeRegion}.
     */
    @Test
    public void testNoShift() {
        final String procName = "testNoShift";
        fileName = procName + ".xlsx";

        testCellAddresses(procName, 0, 0);
        testMergeRegion(procName, 0, 0);
    }

    @Test
    public void testShiftOneRowAndTestAddresses() {
        final String procName = "testShiftOneRowAndTestAddresses";
        fileName = procName + ".xlsx";
        final int nRowsToShift = 1;

        sheet.shiftRows(INSERT_ROW, numRows - 1, nRowsToShift);
        testCellAddresses(procName, nRowsToShift, 0);
    }

    @Test
    public void testShiftOneRowAndTestMergeRegion() {
        final String procName = "testShiftOneRowAndTestMergeRegion";
        fileName = procName + ".xlsx";
        final int nRowsToShift = 1;

        sheet.shiftRows(INSERT_ROW, numRows - 1, nRowsToShift);
        testMergeRegion(procName, nRowsToShift, 0);
    }

    @Test
    public void testShiftOneColumnAndTestAddresses() {
        final String procName = "testShiftOneColumnAndTestAddresses";
        fileName = procName + ".xlsx";
        final int nShift = 1;

        sheet.shiftColumns(INSERT_COLUMN, numCols - 1, nShift);
        testCellAddresses(procName, 0, nShift);
    }

    @Test
    public void testShiftOneColumnAndTestMergeRegion() {
        final String procName = "testShiftOneColumnAndTestMergeRegion";
        fileName = procName + ".xlsx";
        final int nShift = 1;

        sheet.shiftColumns(INSERT_COLUMN, numCols - 1, nShift);
        testMergeRegion(procName, 0, nShift);
    }

    /**
     * Verify that the cell addresses are consistent
     */
    private void testCellAddresses(String procName, int nRowsToShift, int nColsToShift) {
        final int nNumRows = nRowsToShift + numCols;
        final int nNumCols = nColsToShift + numCols;
        for (int nRow = 0; nRow < nNumRows; ++nRow) {
            final XSSFRow row = sheet.getRow(nRow);
            if (row == null) {
                continue;
            }
            for (int nCol = 0; nCol < nNumCols; ++nCol) {
                final String address = new CellAddress(nRow, nCol).formatAsString();
                final XSSFCell cell = row.getCell(nCol);
                if (cell == null) {
                    continue;
                }
                final CTCell ctCell = cell.getCTCell();
                final Object cellAddress = cell.getAddress().formatAsString();
                final Object r = ctCell.getR();

                assertEquals(String.format(Locale.US, "%s: Testing cell.getAddress", procName),
                        address, cellAddress);
                assertEquals(String.format(Locale.US, "%s: Testing ctCell.getR", procName),
                        address, r);
            }
        }

    }

    /**
     * Verify that the merge area is consistent
     */
    private void testMergeRegion(String procName, int nRowsToShift, int nColsToShift) {
        final CellRangeAddress range = sheet.getMergedRegion(0);
        assertEquals(String.format(Locale.US, "%s: Testing merge area %s", procName, range)
                , new CellRangeAddress(FIRST_MERGE_ROW + nRowsToShift, LAST_MERGE_ROW + nRowsToShift,
                        FIRST_MERGE_COL + nColsToShift, LAST_MERGE_COL + nColsToShift), range);
    }
}