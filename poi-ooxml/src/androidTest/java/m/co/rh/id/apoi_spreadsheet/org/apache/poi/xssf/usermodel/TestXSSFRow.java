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
import static org.junit.Assert.assertNotSame;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;
import java.io.OutputStream;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.common.usermodel.HyperlinkType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.tests.usermodel.BaseTestXRow;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellCopyContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellCopyPolicy;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Hyperlink;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.RichTextString;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;

/**
 * Tests for XSSFRow
 */
@RunWith(POIJUnit4ClassRunner.class)
public final class TestXSSFRow extends BaseTestXRow {

    public TestXSSFRow() {
        super(XSSFITestDataProvider.instance);
    }

    @Test
    public void testCopyRowFrom() throws IOException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final XSSFSheet sheet = workbook.createSheet("test");
        final XSSFRow srcRow = sheet.createRow(0);
        srcRow.createCell(0).setCellValue("Hello");
        final XSSFRow destRow = sheet.createRow(1);

        destRow.copyRowFrom(srcRow, new CellCopyPolicy());
        assertNotNull(destRow.getCell(0));
        assertEquals("Hello", destRow.getCell(0).getStringCellValue());

        workbook.close();
    }

    @Test
    public void testCopyRowFromExternalSheet() throws IOException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final Sheet srcSheet = workbook.createSheet("src");
        final XSSFSheet destSheet = workbook.createSheet("dest");
        workbook.createSheet("other");

        final Row srcRow = srcSheet.createRow(0);
        int col = 0;
        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("B5");
        srcRow.createCell(col++).setCellFormula("src!B5");
        srcRow.createCell(col++).setCellFormula("dest!B5");
        srcRow.createCell(col++).setCellFormula("other!B5");

        //Test 2D and 3D Ref Ptgs with absolute row
        srcRow.createCell(col++).setCellFormula("B$5");
        srcRow.createCell(col++).setCellFormula("src!B$5");
        srcRow.createCell(col++).setCellFormula("dest!B$5");
        srcRow.createCell(col++).setCellFormula("other!B$5");

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("SUM(B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(src!B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(dest!B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(other!B5:D$5)");

        //////////////////

        final int styleCount = workbook.getNumCellStyles();

        final XSSFRow destRow = destSheet.createRow(1);
        destRow.copyRowFrom(srcRow, new CellCopyPolicy());

        //////////////////

        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        col = 0;
        Cell cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "src!B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "dest!B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B6", cell.getCellFormula());

        /////////////////////////////////////////////

        //Test 2D and 3D Ref Ptgs with absolute row (Ptg row number shouldn't change)
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "src!B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "dest!B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B$5", cell.getCellFormula());

        //////////////////////////////////////////

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        // Note: absolute row changes from last cell to first cell in order
        // to maintain topLeft:bottomRight order
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area2DPtg",
                "SUM(B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(src!B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(dest!B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(other!B$5:D6)", cell.getCellFormula());

        assertEquals("no new styles should be added by copyRow",
                styleCount, workbook.getNumCellStyles());
        workbook.close();
    }

    @Test
    public void testCopyRowFromDifferentXssfWorkbook() throws IOException {
        final XSSFWorkbook srcWorkbook = new XSSFWorkbook();
        final XSSFWorkbook destWorkbook = new XSSFWorkbook();
        final Sheet srcSheet = srcWorkbook.createSheet("src");
        final XSSFSheet destSheet = destWorkbook.createSheet("dest");
        srcWorkbook.createSheet("other");
        destWorkbook.createSheet("other");

        final Row srcRow = srcSheet.createRow(0);
        int col = 0;
        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("B5");
        srcRow.createCell(col++).setCellFormula("src!B5");
        srcRow.createCell(col++).setCellFormula("dest!B5");
        srcRow.createCell(col++).setCellFormula("other!B5");

        //Test 2D and 3D Ref Ptgs with absolute row
        srcRow.createCell(col++).setCellFormula("B$5");
        srcRow.createCell(col++).setCellFormula("src!B$5");
        srcRow.createCell(col++).setCellFormula("dest!B$5");
        srcRow.createCell(col++).setCellFormula("other!B$5");

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("SUM(B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(src!B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(dest!B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(other!B5:D$5)");

        //////////////////

        final int destStyleCount = destWorkbook.getNumCellStyles();
        final XSSFRow destRow = destSheet.createRow(1);
        destRow.copyRowFrom(srcRow, new CellCopyPolicy(), new CellCopyContext());

        //////////////////

        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        col = 0;
        Cell cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "src!B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "dest!B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B6", cell.getCellFormula());

        /////////////////////////////////////////////

        //Test 2D and 3D Ref Ptgs with absolute row (Ptg row number shouldn't change)
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "src!B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "dest!B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B$5", cell.getCellFormula());

        //////////////////////////////////////////

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        // Note: absolute row changes from last cell to first cell in order
        // to maintain topLeft:bottomRight order
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area2DPtg",
                "SUM(B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(src!B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(dest!B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(other!B$5:D6)", cell.getCellFormula());

        assertEquals("srcWorkbook styles",
                1, srcWorkbook.getNumCellStyles());
        assertEquals("destWorkbook styles",
                destStyleCount + 1, destWorkbook.getNumCellStyles());
        srcWorkbook.close();
        destWorkbook.close();
    }

    @Test
    public void testCopyRowFromDifferentHssfWorkbook() throws IOException {
        final HSSFWorkbook srcWorkbook = new HSSFWorkbook();
        final XSSFWorkbook destWorkbook = new XSSFWorkbook();
        final Sheet srcSheet = srcWorkbook.createSheet("src");
        final XSSFSheet destSheet = destWorkbook.createSheet("dest");
        srcWorkbook.createSheet("other");
        destWorkbook.createSheet("other");

        final Row srcRow = srcSheet.createRow(0);
        int col = 0;
        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("B5");
        srcRow.createCell(col++).setCellFormula("other!B5");

        //Test 2D and 3D Ref Ptgs with absolute row
        srcRow.createCell(col++).setCellFormula("B$5");
        srcRow.createCell(col++).setCellFormula("other!B$5");

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        srcRow.createCell(col++).setCellFormula("SUM(B5:D$5)");
        srcRow.createCell(col++).setCellFormula("SUM(other!B5:D$5)");

        //////////////////

        final int destStyleCount = destWorkbook.getNumCellStyles();
        final XSSFRow destRow = destSheet.createRow(1);
        CellCopyPolicy policy = new CellCopyPolicy();
        //hssf to xssf copy does not support cell style copying
        policy.setCopyCellStyle(false);
        destRow.copyRowFrom(srcRow, policy, new CellCopyContext());

        //////////////////

        //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
        col = 0;
        Cell cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B6", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B6", cell.getCellFormula());

        /////////////////////////////////////////////

        //Test 2D and 3D Ref Ptgs with absolute row (Ptg row number shouldn't change)
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("RefPtg",
                "B$5", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Ref3DPtg",
                "other!B$5", cell.getCellFormula());

        //////////////////////////////////////////

        //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
        // Note: absolute row changes from last cell to first cell in order
        // to maintain topLeft:bottomRight order
        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area2DPtg",
                "SUM(B$5:D6)", cell.getCellFormula());

        cell = destRow.getCell(col++);
        assertNotNull(cell);
        assertEquals("Area3DPtg",
                "SUM(other!B$5:D6)", cell.getCellFormula());

        assertEquals("destWorkbook styles",
                destStyleCount, destWorkbook.getNumCellStyles());
        srcWorkbook.close();
        destWorkbook.close();
    }

    @Test
    public void testCopyRowWithHyperlink() throws IOException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final Sheet srcSheet = workbook.createSheet("src");
        final XSSFSheet destSheet = workbook.createSheet("dest");
        Row srcRow = srcSheet.createRow(0);
        Cell srcCell = srcRow.createCell(0);
        Hyperlink srcHyperlink = new XSSFHyperlink(HyperlinkType.URL);
        srcHyperlink.setAddress("https://poi.apache.org");
        srcCell.setHyperlink(srcHyperlink);

        final XSSFRow destRow = destSheet.createRow(0);
        destRow.copyRowFrom(srcRow, new CellCopyPolicy());

        Cell destCell = destRow.getCell(0);
        assertEquals("cell hyperlink addresses match",
                srcCell.getHyperlink().getAddress(), destCell.getHyperlink().getAddress());
        assertEquals("cell hyperlink types match",
                srcCell.getHyperlink().getType(), destCell.getHyperlink().getType());

        workbook.close();
    }

    @Test
    public void testCopyRowOverwritesExistingRow() throws IOException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final XSSFSheet sheet1 = workbook.createSheet("Sheet1");
        final Sheet sheet2 = workbook.createSheet("Sheet2");

        final Row srcRow = sheet1.createRow(0);
        final XSSFRow destRow = sheet1.createRow(1);
        final Row observerRow = sheet1.createRow(2);
        final Row externObserverRow = sheet2.createRow(0);

        srcRow.createCell(0).setCellValue("hello");
        srcRow.createCell(1).setCellValue("world");
        destRow.createCell(0).setCellValue(5.0); //A2 -> 5.0
        destRow.createCell(1).setCellFormula("A1"); // B2 -> A1 -> "hello"
        observerRow.createCell(0).setCellFormula("A2"); // A3 -> A2 -> 5.0
        observerRow.createCell(1).setCellFormula("B2"); // B3 -> B2 -> A1 -> "hello"
        externObserverRow.createCell(0).setCellFormula("Sheet1!A2"); //Sheet2!A1 -> Sheet1!A2 -> 5.0

        // overwrite existing destRow with row-copy of srcRow
        destRow.copyRowFrom(srcRow, new CellCopyPolicy());

        // copyRowFrom should update existing destRow, rather than creating a new row and reassigning the destRow pointer
        // to the new row (and allow the old row to be garbage collected)
        // this is mostly so existing references to rows that are overwritten are updated
        // rather than allowing users to continue updating rows that are no longer part of the sheet
        assertSame("existing references to srcRow are still valid",
                srcRow, sheet1.getRow(0));
        assertSame("existing references to destRow are still valid",
                destRow, sheet1.getRow(1));
        assertSame("existing references to observerRow are still valid",
                observerRow, sheet1.getRow(2));
        assertSame("existing references to externObserverRow are still valid",
                externObserverRow, sheet2.getRow(0));

        // Make sure copyRowFrom actually copied row (this is tested elsewhere)
        assertEquals(CellType.STRING, destRow.getCell(0).getCellType());
        assertEquals("hello", destRow.getCell(0).getStringCellValue());

        // We don't want #REF! errors if we copy a row that contains cells that are referred to by other cells outside of copied region
        assertEquals("references to overwritten cells are unmodified",
                "A2", observerRow.getCell(0).getCellFormula());
        assertEquals("references to overwritten cells are unmodified",
                "B2", observerRow.getCell(1).getCellFormula());
        assertEquals("references to overwritten cells are unmodified",
                "Sheet1!A2", externObserverRow.getCell(0).getCellFormula());

        workbook.close();
    }

    @Test
    public void testMultipleEditWriteCycles() {
        final XSSFWorkbook wb1 = new XSSFWorkbook();
        final XSSFSheet sheet1 = wb1.createSheet("Sheet1");
        XSSFRow srcRow = sheet1.createRow(0);
        srcRow.createCell(0).setCellValue("hello");
        srcRow.createCell(3).setCellValue("world");

        // discard result
        XSSFTestDataSamples.writeOutAndReadBack(wb1);

        assertEquals("hello", srcRow.getCell(0).getStringCellValue());
        assertEquals("hello",
                wb1.getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue());
        assertEquals("world", srcRow.getCell(3).getStringCellValue());
        assertEquals("world",
                wb1.getSheet("Sheet1").getRow(0).getCell(3).getStringCellValue());

        srcRow.createCell(1).setCellValue("cruel");

        // discard result
        XSSFTestDataSamples.writeOutAndReadBack(wb1);

        assertEquals("hello", srcRow.getCell(0).getStringCellValue());
        assertEquals("hello", wb1.getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue());
        assertEquals("cruel", srcRow.getCell(1).getStringCellValue());
        assertEquals("cruel", wb1.getSheet("Sheet1").getRow(0).getCell(1).getStringCellValue());
        assertEquals("world", srcRow.getCell(3).getStringCellValue());
        assertEquals("world", wb1.getSheet("Sheet1").getRow(0).getCell(3).getStringCellValue());

        srcRow.getCell(1).setCellValue((RichTextString) null);

        XSSFWorkbook wb3 = XSSFTestDataSamples.writeOutAndReadBack(wb1);
        assertEquals("Cell should be blank",
                CellType.BLANK, wb3.getSheet("Sheet1").getRow(0).getCell(1).getCellType());
    }

    @Test
    public void createHeaderOnce() throws IOException {
        Workbook wb = new XSSFWorkbook();
        fillData(1, wb.createSheet("sheet123"));
        writeToFile(wb);
    }

    @Test
    public void createHeaderTwice() throws IOException {
        Workbook wb = new XSSFWorkbook();
        fillData(0, wb.createSheet("sheet123"));
        writeToFile(wb);
    }

    @Test
    public void createHeaderThreeTimes() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet123");
        fillData(1, sheet);
        fillData(0, sheet);
        fillData(0, sheet);
        writeToFile(wb);
    }

    @Test
    public void duplicateRows() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("sheet123");
            // it is not a good idea to create a row twice but does it shouldn't fail
            // you would likely lose all the cells associated with the first row instance
            // ie when you write the file you will only have 1 row2 and only the cells for the 2nd row instance
            XSSFRow rowX = sheet.createRow(2);
            rowX.createCell(0).setCellValue("rowX-c0");
            XSSFRow rowY = sheet.createRow(2);
            rowY.createCell(1).setCellValue("rowY-c1");
            assertNotSame(rowX, rowY);

            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                wb.write(bos);
                try (XSSFWorkbook wb2 = new XSSFWorkbook(bos.toInputStream())) {
                    XSSFSheet sheet2 = wb2.getSheet(sheet.getSheetName());
                    XSSFRow rowZ = sheet2.getRow(2);
                    assertNull(rowZ.getCell(0));
                    assertEquals(rowY.getCell(1).getStringCellValue(), rowZ.getCell(1).getStringCellValue());
                }
            }
        }
    }

    private void fillData(int startAtRow, Sheet sheet) {
        Row header = sheet.createRow(0);
        for (int rownum = startAtRow; rownum < 2; rownum++) {
            header.createCell(0).setCellValue("a");

            Row row = sheet.createRow(rownum);
            row.createCell(0).setCellValue("a");
        }
    }

    private void writeToFile(Workbook wb) throws IOException {
        try (OutputStream fileOut = UnsynchronizedByteArrayOutputStream.builder().get()) {
            wb.write(fileOut);
        }
    }
}