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
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertThrows;
import static org.junit.Assert.assertTrue;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.ITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BaseTestBugzillaIssues;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CreationHelper;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FillPatternType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Font;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.HorizontalAlignment;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.IndexedColors;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.PrintSetup;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.SXSSFITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFCell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFRow;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public final class TestSXSSFBugs extends BaseTestBugzillaIssues {
    public TestSXSSFBugs() {
        super(SXSSFITestDataProvider.instance);
    }

    // override some tests which do not work for SXSSF
    @Override
    @Ignore("cloneSheet() not implemented")
    public void bug18800() { /* cloneSheet() not implemented */ }

    @Override
    @Ignore("cloneSheet() not implemented")
    public void bug22720() { /* cloneSheet() not implemented */ }

    @Override
    @Ignore("Evaluation is not fully supported")
    public void bug47815() { /* Evaluation is not supported */ }

    @Override
    @Ignore("Evaluation is not fully supported")
    public void bug46729_testMaxFunctionArguments() { /* Evaluation is not supported */ }

    @Override
    @Ignore("Reading data is not supported")
    public void bug57798() { /* Reading data is not supported */ }

    /**
     * Setting repeating rows and columns shouldn't break
     * any print settings that were there before
     */
    @Test
    public void bug49253() throws Exception {
        Workbook wb1 = new SXSSFWorkbook();
        Workbook wb2 = new SXSSFWorkbook();
        CellRangeAddress cra = CellRangeAddress.valueOf("C2:D3");

        // No print settings before repeating
        Sheet s1 = wb1.createSheet();
        s1.setRepeatingColumns(cra);
        s1.setRepeatingRows(cra);

        PrintSetup ps1 = s1.getPrintSetup();
        assertFalse(ps1.getValidSettings());
        assertFalse(ps1.getLandscape());


        // Had valid print settings before repeating
        Sheet s2 = wb2.createSheet();
        PrintSetup ps2 = s2.getPrintSetup();

        ps2.setLandscape(false);
        assertTrue(ps2.getValidSettings());
        assertFalse(ps2.getLandscape());
        s2.setRepeatingColumns(cra);
        s2.setRepeatingRows(cra);

        ps2 = s2.getPrintSetup();
        assertTrue(ps2.getValidSettings());
        assertFalse(ps2.getLandscape());

        wb1.close();
        wb2.close();
    }

    @Test
    public void bug61648() throws Exception {
        // works as expected
        writeWorkbook(new XSSFWorkbook(), XSSFITestDataProvider.instance);

        // does not work
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            assertThrows("this is not implemented yet",
                    RuntimeException.class, () -> writeWorkbook(wb, SXSSFITestDataProvider.instance));
        }
    }

    void writeWorkbook(Workbook wb, ITestDataProvider testDataProvider) throws IOException {
        Sheet sheet = wb.createSheet("array formula test");

        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex++);

        Cell cell = row.createCell(colIndex++);
        cell.setCellValue("multiple");
        cell = row.createCell(colIndex);
        cell.setCellValue("unique");

        writeRow(sheet, rowIndex++, 80d, "INDEX(A2:A7, MATCH(FALSE, ISBLANK(A2:A7), 0))");
        writeRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B2, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
        writeRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B3, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
        writeRow(sheet, rowIndex++, 2d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B4, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
        writeRow(sheet, rowIndex++, 30d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B5, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");
        writeRow(sheet, rowIndex, 2d, "IFERROR(INDEX(A2:A7, MATCH(1, (COUNTIF(B2:B6, A2:A7) = 0) * (NOT(ISBLANK(A2:A7))), 0)), \"\")");

        /*FileOutputStream fileOut = new FileOutputStream(filename);
        wb.write(fileOut);
        fileOut.close();*/

        Workbook wbBack = testDataProvider.writeOutAndReadBack(wb);
        assertNotNull(wbBack);
        wbBack.close();

        wb.close();
    }

    void writeRow(Sheet sheet, int rowIndex, Double col0Value, String col1Value) {
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex);

        // numeric value cell
        Cell cell = row.createCell(colIndex++);
        cell.setCellValue(col0Value);

        // formula value cell
        CellRangeAddress range = new CellRangeAddress(rowIndex, rowIndex, colIndex, colIndex);
        sheet.setArrayFormula(col1Value, range);
    }

    @Ignore("takes too long for the normal test run")
    public void test62872() throws Exception {
        final int COLUMN_COUNT = 300;
        final int ROW_COUNT = 600000;
        final int TEN_MINUTES = 1000 * 60 * 10;

        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        workbook.setCompressTempFiles(true);
        SXSSFSheet sheet = workbook.createSheet("RawData");

        SXSSFRow row = sheet.createRow(0);
        SXSSFCell cell;

        for (int i = 1; i <= COLUMN_COUNT; i++) {
            cell = row.createCell(i - 1);
            cell.setCellValue("Column " + i);
        }

        for (int i = 1; i < ROW_COUNT; i++) {
            row = sheet.createRow(i);
            for (int j = 1; j <= COLUMN_COUNT; j++) {
                cell = row.createCell(j - 1);

                //make some noise
                cell.setCellValue(new Date(i * TEN_MINUTES + (j * TEN_MINUTES) / COLUMN_COUNT));
            }
            i++;
        }

        try (FileOutputStream out = new FileOutputStream(File.createTempFile("test62872", ".xlsx"))) {
            workbook.write(out);
            workbook.dispose();
            workbook.close();
            out.flush();
        }
    }

    @Test
    public void test63960() throws Exception {
        try (Workbook workbook = new SXSSFWorkbook(100)) {
            Sheet sheet = workbook.createSheet("RawData");

            Row row = sheet.createRow(0);
            Cell cell;

            cell = row.createCell(0);
            cell.setCellValue(123);
            cell = row.createCell(1);
            cell.setCellValue("");
            cell.setCellFormula("=TEXT(A1,\"#\")");

            /*try (FileOutputStream out = new FileOutputStream(File.createTempFile("test63960", ".xlsx"))) {
                workbook.write(out);
            }*/

            try (Workbook wbBack = SXSSFITestDataProvider.instance.writeOutAndReadBack(workbook)) {
                assertNotNull(wbBack);
                Cell rawData = wbBack.getSheet("RawData").getRow(0).getCell(1);

                assertEquals(STCellType.STR, ((XSSFCell) rawData).getCTCell().getT());
            }
        }
    }

    @Test
    public void test64595() throws Exception {
        try (Workbook workbook = new SXSSFWorkbook(100)) {
            Sheet sheet = workbook.createSheet("RawData");
            Row row = sheet.createRow(0);
            Cell cell;

            cell = row.createCell(0);
            cell.setCellValue("Ernie & Bert");

            cell = row.createCell(1);
            // Set a precalculated formula value containing a special character.
            cell.setCellValue("Ernie & Bert are cool!");
            cell.setCellFormula("A1 & \" are cool!\"");

            // While unfixed reading the workbook would throw a POIXMLException
            // since the file was corrupt due to missing quotation.
            try (Workbook wbBack = SXSSFITestDataProvider.instance.writeOutAndReadBack(workbook)) {
                assertNotNull(wbBack);
                cell = wbBack.getSheetAt(0).getRow(0).getCell(1);
                assertEquals("Ernie & Bert are cool!", cell.getStringCellValue());
                assertEquals("A1 & \" are cool!\"", cell.getCellFormula());
            }
        }
    }

    @Test
    public void test65619() throws Exception {
        try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
            try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
                SXSSFSheet sheet = workbook.createSheet("Test Sheet 1");
                Font font = workbook.createFont();
                font.setItalic(true);
                CellStyle cs = workbook.createCellStyle();
                cs.setFont(font);
                Row row = sheet.createRow(0);
                Cell cell = row.createCell(1);
                cell.setCellValue(new Date());
                cell.setCellStyle(cs);
                CreationHelper ch = workbook.getCreationHelper();
                cs.setDataFormat(ch.createDataFormat().getFormat("dd MMM yyyy HH:mm:ss"));
                cs.setAlignment(HorizontalAlignment.RIGHT);
                workbook.write(bos);
            }
            try (XSSFWorkbook workbook = new XSSFWorkbook(bos.toInputStream())) {
                XSSFSheet sheet = workbook.getSheet("Test Sheet 1");
                XSSFRow row = sheet.getRow(0);
                XSSFCell cell = row.getCell(1);
                XSSFCellStyle cs = cell.getCellStyle();
                assertNotNull("cell should have style", cs);
                assertEquals(HorizontalAlignment.RIGHT, cs.getCellAlignment().getHorizontal());
                assertEquals("dd MMM yyyy HH:mm:ss", cs.getDataFormatString());
                XSSFFont font = cs.getFont();
                assertNotNull("style should have font", font);
                assertTrue("saved font is italic", font.getItalic());
            }
        }
    }

    @Test
    public void test62215() throws IOException {
        String sheetName = "1";
        int rowIndex = 0;
        int colIndex = 0;
        String formula = "1";
        String value = "yes";
        CellType valueType = CellType.STRING;

        try (Workbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet(sheetName);
            Row row = sheet.createRow(rowIndex);
            Cell cell = row.createCell(colIndex);

            // this order ensures that value will not be overwritten by setting the formula
            cell.setCellFormula(formula);
            cell.setCellValue(value);

            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals(formula, cell.getCellFormula());
            assertEquals(valueType, cell.getCachedFormulaResultType());
            assertEquals(value, cell.getStringCellValue());
            // so far so good

            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                wb.write(bos);

                try (XSSFWorkbook testWb = new XSSFWorkbook(bos.toInputStream())) {
                    Cell testCell = testWb.getSheet(sheetName).getRow(rowIndex).getCell(colIndex);
                    assertEquals(CellType.FORMULA, testCell.getCellType());
                    assertEquals(formula, testCell.getCellFormula());

                    assertEquals(CellType.STRING, testCell.getCachedFormulaResultType());
                    assertEquals(value, testCell.getStringCellValue());
                }
            }
        }
    }

    @Test
    public void testBug51037() throws IOException {
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            CellStyle blueStyle = wb.createCellStyle();
            blueStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            blueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle pinkStyle = wb.createCellStyle();
            pinkStyle.setFillForegroundColor(IndexedColors.PINK.getIndex());
            pinkStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Sheet s1 = wb.createSheet("Pretty columns");

            s1.setDefaultColumnStyle(4, blueStyle);
            s1.setDefaultColumnStyle(6, pinkStyle);

            Row r3 = s1.createRow(3);
            r3.createCell(0).setCellValue("The");
            r3.createCell(1).setCellValue("quick");
            r3.createCell(2).setCellValue("brown");
            r3.createCell(3).setCellValue("fox");
            r3.createCell(4).setCellValue("jumps");
            r3.createCell(5).setCellValue("over");
            r3.createCell(6).setCellValue("the");
            r3.createCell(7).setCellValue("lazy");
            r3.createCell(8).setCellValue("dog");
            Row r7 = s1.createRow(7);
            r7.createCell(1).setCellStyle(pinkStyle);
            r7.createCell(8).setCellStyle(blueStyle);

            assertEquals(blueStyle.getIndex(), r3.getCell(4).getCellStyle().getIndex());
            assertEquals(pinkStyle.getIndex(), r3.getCell(6).getCellStyle().getIndex());

            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                wb.write(bos);
                try (XSSFWorkbook wb2 = new XSSFWorkbook(bos.toInputStream())) {
                    XSSFSheet wb2Sheet = wb2.getSheetAt(0);
                    XSSFRow wb2R3 = wb2Sheet.getRow(3);
                    assertEquals(blueStyle.getIndex(), wb2R3.getCell(4).getCellStyle().getIndex());
                    assertEquals(pinkStyle.getIndex(), wb2R3.getCell(6).getCellStyle().getIndex());
                }
            }
        }
    }
}
