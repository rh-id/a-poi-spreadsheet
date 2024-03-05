/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertThrows;
import static org.mockito.Mockito.spy;

import org.apache.xmlbeans.XmlCursor;
import org.junit.AfterClass;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.IOException;
import java.nio.charset.StandardCharsets;

import javax.xml.namespace.QName;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.SpreadsheetVersion;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.tests.usermodel.BaseTestXCell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.RichTextString;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.SXSSFITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFCell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFRichTextString;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Tests various functionality having to do with {@link SXSSFCell}.  For instance support for
 * particular datatypes, etc.
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestSXSSFCell extends BaseTestXCell {

    public TestSXSSFCell() {
        super(SXSSFITestDataProvider.instance);
    }

    @AfterClass
    public static void tearDownClass() {
        SXSSFITestDataProvider.instance.cleanup();
    }

    @Test
    public void testPreserveSpaces() throws IOException {
        String[] samplesWithSpaces = {
                " POI",
                "POI ",
                " POI ",
                "\nPOI",
                "\n\nPOI \n",
        };
        for (String str : samplesWithSpaces) {
            try (Workbook swb = _testDataProvider.createWorkbook()) {
                Cell sCell = swb.createSheet().createRow(0).createCell(0);
                sCell.setCellValue(str);
                assertEquals(sCell.getStringCellValue(), str);

                // read back as XSSF and check that xml:spaces="preserve" is set
                try (XSSFWorkbook xwb = (XSSFWorkbook) _testDataProvider.writeOutAndReadBack(swb)) {
                    XSSFCell xCell = xwb.getSheetAt(0).getRow(0).getCell(0);

                    CTRst is = xCell.getCTCell().getIs();
                    assertNotNull(is);
                    try (XmlCursor c = is.newCursor()) {
                        c.toNextToken();
                        String t = c.getAttributeText(new QName("http://www.w3.org/XML/1998/namespace", "space"));
                        assertEquals("expected xml:spaces=\"preserve\" \"" + str + "\"",
                                "preserve", t);
                    }
                }
            }
        }
    }

    @Test
    public void getCachedFormulaResultType_throwsISE_whenNotAFormulaCell() {
        SXSSFCell instance = new SXSSFCell(null, CellType.BLANK);
        assertThrows(IllegalStateException.class, instance::getCachedFormulaResultType);
    }


    @Test
    public void setCellValue_withTooLongRichTextString_throwsIAE() {
        Cell cell = spy(new SXSSFCell(null, CellType.BLANK));
        int length = SpreadsheetVersion.EXCEL2007.getMaxTextLength() + 1;
        String string = new String(new byte[length], StandardCharsets.UTF_8).replace("\0", "x");
        RichTextString richTextString = new XSSFRichTextString(string);
        assertThrows(IllegalArgumentException.class, () -> cell.setCellValue(richTextString));
    }

    @Test
    public void getArrayFormulaRange_returnsNull() {
        Cell cell = new SXSSFCell(null, CellType.BLANK);
        CellRangeAddress result = cell.getArrayFormulaRange();
        assertNull(result);
    }

    @Test
    public void isPartOfArrayFormulaGroup_returnsFalse() {
        Cell cell = new SXSSFCell(null, CellType.BLANK);
        boolean result = cell.isPartOfArrayFormulaGroup();
        assertFalse(result);
    }

    @Test
    public void getErrorCellValue_returns0_onABlankCell() {
        Cell cell = new SXSSFCell(null, CellType.BLANK);
        assertEquals(CellType.BLANK, cell.getCellType());
        byte result = cell.getErrorCellValue();
        assertEquals(0, result);
    }

    /**
     * For now, {@link SXSSFCell} doesn't support array formulas.
     * However, this test should be enabled if array formulas are implemented for SXSSF.
     */
    @Override
    @Ignore
    public void setBlank_removesArrayFormula_ifCellIsPartOfAnArrayFormulaGroupContainingOnlyThisCell() {
    }

    /**
     * For now, {@link SXSSFCell} doesn't support array formulas.
     * However, this test should be enabled if array formulas are implemented for SXSSF.
     */
    @Override
    @Ignore
    public void setBlank_throwsISE_ifCellIsPartOfAnArrayFormulaGroupContainingOtherCells() {
    }

    @Override
    @Ignore
    public void setCellFormula_throwsISE_ifCellIsPartOfAnArrayFormulaGroupContainingOtherCells() {
    }

    @Override
    @Ignore
    public void removeFormula_turnsCellToBlank_whenFormulaWasASingleCellArrayFormula() {
    }

    @Override
    @Ignore
    public void setCellFormula_onASingleCellArrayFormulaCell_preservesTheValue() {
    }

    @Override
    @Ignore
    public void setCellType_FORMULA_onAnArrayFormulaCell_doesNothing() {
    }
}
