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

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNotSame;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertThrows;
import static org.junit.Assert.assertTrue;
import static m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontCharset.ANSI;
import static m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontCharset.ARABIC;
import static m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontCharset.DEFAULT;
import static m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontCharset.RUSSIAN;
import static m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontCharset.valueOf;

import android.graphics.Color;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontFamily;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontSize;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIntProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTVerticalAlignFontProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFontScheme;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.common.usermodel.fonts.FontCharset;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ooxml.POIXMLException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BaseTestFont;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Font;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontFamily;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontScheme;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontUnderline;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.IndexedColors;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.SheetUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LocaleUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;

@RunWith(POIJUnit4ClassRunner.class)
public final class TestXSSFFont extends BaseTestFont {

    public TestXSSFFont() {
        super(XSSFITestDataProvider.instance);
    }

    @Override
    protected Iterable<Object[]> defaultFont() {
        List<Object[]> data = new ArrayList<>();
        data.add(new Object[]{"Calibri", (short) 220, IndexedColors.BLACK.getIndex()});
        return data;
    }

    @Test
    public void testConstructor() {
        XSSFFont xssfFont = new XSSFFont();
        assertNotNull(xssfFont.getCTFont());
    }

    @Test
    public void testBold() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTBooleanProperty bool = ctFont.addNewB();
        bool.setVal(false);
        ctFont.setBArray(0, bool);
        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertFalse(xssfFont.getBold());


        xssfFont.setBold(true);
        assertEquals(1, ctFont.sizeOfBArray());
        assertTrue(ctFont.getBArray(0).getVal());
    }

    @SuppressWarnings("deprecation")
    @Test
    public void testCharSetWithDeprecatedFontCharset() throws IOException {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTIntProperty prop = ctFont.addNewCharset();
        prop.setVal(ANSI.getValue());

        ctFont.setCharsetArray(0, prop);
        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(Font.ANSI_CHARSET, xssfFont.getCharSet());

        xssfFont.setCharSet(DEFAULT);
        assertEquals(DEFAULT.getValue(), ctFont.getCharsetArray(0).getVal());

        // Try with a few less usual ones:
        // Set with the Charset itself
        xssfFont.setCharSet(RUSSIAN);
        assertEquals(RUSSIAN.getValue(), xssfFont.getCharSet());
        // And set with the Charset index
        xssfFont.setCharSet(ARABIC.getValue());
        assertEquals(ARABIC.getValue(), xssfFont.getCharSet());
        xssfFont.setCharSet((byte) (ARABIC.getValue()));
        assertEquals(ARABIC.getValue(), xssfFont.getCharSet());

        // This one isn't allowed
        assertNull(valueOf(9999));
        assertThrows("Shouldn't be able to set an invalid charset",
                POIXMLException.class, () -> xssfFont.setCharSet(9999));

        // Now try with a few sample files

        // Normal charset
        try (XSSFWorkbook wb1 = XSSFTestDataSamples.openSampleWorkbook("Formatting.xlsx")) {
            assertEquals(0,
                    wb1.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFont().getCharSet()
            );
        }

        // GB2312 charset
        try (XSSFWorkbook wb2 = XSSFTestDataSamples.openSampleWorkbook("49273.xlsx")) {
            assertEquals(134,
                    wb2.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFont().getCharSet()
            );
        }
    }

    @Test
    public void testCharSet() throws IOException {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTIntProperty prop = ctFont.addNewCharset();
        prop.setVal(FontCharset.ANSI.getNativeId());

        ctFont.setCharsetArray(0, prop);
        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(Font.ANSI_CHARSET, xssfFont.getCharSet());

        xssfFont.setCharSet(FontCharset.DEFAULT);
        assertEquals(FontCharset.DEFAULT.getNativeId(), ctFont.getCharsetArray(0).getVal());

        // Try with a few less usual ones:
        // Set with the Charset itself
        xssfFont.setCharSet(FontCharset.RUSSIAN);
        assertEquals(FontCharset.RUSSIAN.getNativeId(), xssfFont.getCharSet());
        // And set with the Charset index
        xssfFont.setCharSet(FontCharset.ARABIC.getNativeId());
        assertEquals(FontCharset.ARABIC.getNativeId(), xssfFont.getCharSet());
        xssfFont.setCharSet((byte) (FontCharset.ARABIC.getNativeId()));
        assertEquals(FontCharset.ARABIC.getNativeId(), xssfFont.getCharSet());

        // This one isn't allowed
        assertNull(FontCharset.valueOf(9999));
        assertThrows("Shouldn't be able to set an invalid charset",
                POIXMLException.class, () -> xssfFont.setCharSet(9999));

        // Now try with a few sample files

        // Normal charset
        try (XSSFWorkbook wb1 = XSSFTestDataSamples.openSampleWorkbook("Formatting.xlsx")) {
            assertEquals(0,
                    wb1.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFont().getCharSet()
            );
        }

        // GB2312 charset
        try (XSSFWorkbook wb2 = XSSFTestDataSamples.openSampleWorkbook("49273.xlsx")) {
            assertEquals(134,
                    wb2.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFont().getCharSet()
            );
        }
    }

    @Test
    public void testFontName() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTFontName fname = ctFont.addNewName();
        fname.setVal("Arial");
        ctFont.setNameArray(0, fname);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals("Arial", xssfFont.getFontName());

        xssfFont.setFontName("Courier");
        assertEquals("Courier", ctFont.getNameArray(0).getVal());
    }

    @Test
    public void testItalic() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTBooleanProperty bool = ctFont.addNewI();
        bool.setVal(false);
        ctFont.setIArray(0, bool);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertFalse(xssfFont.getItalic());

        xssfFont.setItalic(true);
        assertEquals(1, ctFont.sizeOfIArray());
        assertTrue(ctFont.getIArray(0).getVal());
        assertTrue(ctFont.getIArray(0).getVal());
    }

    @Test
    public void testStrikeout() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTBooleanProperty bool = ctFont.addNewStrike();
        bool.setVal(false);
        ctFont.setStrikeArray(0, bool);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertFalse(xssfFont.getStrikeout());

        xssfFont.setStrikeout(true);
        assertEquals(1, ctFont.sizeOfStrikeArray());
        assertTrue(ctFont.getStrikeArray(0).getVal());
        assertTrue(ctFont.getStrikeArray(0).getVal());
    }

    @Test
    public void testFontHeight() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTFontSize size = ctFont.addNewSz();
        size.setVal(11);
        ctFont.setSzArray(0, size);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(11, xssfFont.getFontHeightInPoints());

        xssfFont.setFontHeight(20);
        assertEquals(20.0, ctFont.getSzArray(0).getVal(), 0.0);
    }

    @Test
    public void testFontHeightInPoint() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTFontSize size = ctFont.addNewSz();
        size.setVal(14);
        ctFont.setSzArray(0, size);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(14, xssfFont.getFontHeightInPoints());

        xssfFont.setFontHeightInPoints((short) 20);
        assertEquals(20.0, ctFont.getSzArray(0).getVal(), 0.0);
    }

    @Test
    public void testUnderline() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTUnderlineProperty underlinePropr = ctFont.addNewU();
        underlinePropr.setVal(STUnderlineValues.SINGLE);
        ctFont.setUArray(0, underlinePropr);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(Font.U_SINGLE, xssfFont.getUnderline());

        xssfFont.setUnderline(Font.U_DOUBLE);
        assertEquals(1, ctFont.sizeOfUArray());
        assertSame(STUnderlineValues.DOUBLE, ctFont.getUArray(0).getVal());

        xssfFont.setUnderline(FontUnderline.DOUBLE_ACCOUNTING);
        assertEquals(1, ctFont.sizeOfUArray());
        assertSame(STUnderlineValues.DOUBLE_ACCOUNTING, ctFont.getUArray(0).getVal());
    }

    @Test
    public void testColor() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTColor color = ctFont.addNewColor();
        color.setIndexed(XSSFFont.DEFAULT_FONT_COLOR);
        ctFont.setColorArray(0, color);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(IndexedColors.BLACK.getIndex(), xssfFont.getColor());

        xssfFont.setColor(IndexedColors.RED.getIndex());
        assertEquals(IndexedColors.RED.getIndex(), ctFont.getColorArray(0).getIndexed());
    }

    @Test
    public void testRgbColor() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTColor color = ctFont.addNewColor();

        color.setRgb(Integer.toHexString(0xFFFFFF).getBytes(LocaleUtil.CHARSET_1252));
        ctFont.setColorArray(0, color);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(ctFont.getColorArray(0).getRgb()[0], xssfFont.getXSSFColor().getRGB()[0]);
        assertEquals(ctFont.getColorArray(0).getRgb()[1], xssfFont.getXSSFColor().getRGB()[1]);
        assertEquals(ctFont.getColorArray(0).getRgb()[2], xssfFont.getXSSFColor().getRGB()[2]);
        assertEquals(ctFont.getColorArray(0).getRgb()[3], xssfFont.getXSSFColor().getRGB()[3]);

        xssfFont.setColor((short) 23);

        byte[] bytes = Integer.toHexString(0xF1F1F1).getBytes(LocaleUtil.CHARSET_1252);
        color.setRgb(bytes);
        XSSFColor newColor = XSSFColor.from(color);
        xssfFont.setColor(newColor);
        assertEquals(ctFont.getColorArray(0).getRgb()[2], newColor.getRGB()[2]);

        assertArrayEquals(bytes, xssfFont.getXSSFColor().getRGB());
        assertEquals(0, xssfFont.getColor());
    }

    @Test
    public void testThemeColor() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTColor color = ctFont.addNewColor();
        color.setTheme(1);
        ctFont.setColorArray(0, color);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(ctFont.getColorArray(0).getTheme(), xssfFont.getThemeColor());

        xssfFont.setThemeColor(IndexedColors.RED.getIndex());
        assertEquals(IndexedColors.RED.getIndex(), ctFont.getColorArray(0).getTheme());
    }

    @Test
    public void testFamily() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTFontFamily family = ctFont.addNewFamily();
        family.setVal(FontFamily.MODERN.getValue());
        ctFont.setFamilyArray(0, family);

        XSSFFont xssfFont = new XSSFFont(ctFont);
        assertEquals(FontFamily.MODERN.getValue(), xssfFont.getFamily());
    }

    @Test
    public void testScheme() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTFontScheme scheme = ctFont.addNewScheme();
        scheme.setVal(STFontScheme.MAJOR);
        ctFont.setSchemeArray(0, scheme);

        XSSFFont font = new XSSFFont(ctFont);
        assertEquals(FontScheme.MAJOR, font.getScheme());

        font.setScheme(FontScheme.NONE);
        assertSame(STFontScheme.NONE, ctFont.getSchemeArray(0).getVal());
    }

    @Test
    public void testTypeOffset() {
        CTFont ctFont = CTFont.Factory.newInstance();
        CTVerticalAlignFontProperty valign = ctFont.addNewVertAlign();
        valign.setVal(STVerticalAlignRun.BASELINE);
        ctFont.setVertAlignArray(0, valign);

        XSSFFont font = new XSSFFont(ctFont);
        assertEquals(Font.SS_NONE, font.getTypeOffset());

        font.setTypeOffset(XSSFFont.SS_SUPER);
        assertSame(STVerticalAlignRun.SUPERSCRIPT, ctFont.getVertAlignArray(0).getVal());
    }

    // store test from TestSheetUtil here as it uses XSSF
    @Test
    public void testCanComputeWidthXSSF() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            // cannot check on result because on some machines we get back false here!
            //assertDoesNotThrow(() -> SheetUtil.canComputeColumnWidth(wb.getFontAt(0)));
            SheetUtil.canComputeColumnWidth(wb.getFontAt(0));
        }
    }

    // store test from TestSheetUtil here as it uses XSSF
    @Test
    public void testCanComputeWidthInvalidFont() {
        Font font = new XSSFFont(CTFont.Factory.newInstance());
        font.setFontName("some non existing font name");

        // Even with invalid fonts we still get back useful data most of the time...
        SheetUtil.canComputeColumnWidth(font);
    }

    /**
     * Test that fonts get added properly
     */
    @Test
    public void testFindFont() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            assertEquals(1, wb.getNumberOfFonts());

            XSSFSheet s = wb.createSheet();
            s.createRow(0);
            s.createRow(1);
            s.getRow(0).createCell(0);
            s.getRow(1).createCell(0);

            assertEquals(1, wb.getNumberOfFonts());

            XSSFFont f1 = wb.getFontAt(0);
            assertFalse(f1.getBold());

            // Check that asking for the same font
            //  multiple times gives you the same thing.
            // Otherwise, our tests wouldn't work!
            assertSame(wb.getFontAt(0), wb.getFontAt(0));
            assertEquals(
                    wb.getFontAt(0),
                    wb.getFontAt(0)
            );

            // Look for a new font we have
            //  yet to add
            assertNull(
                    wb.findFont(
                            false, IndexedColors.INDIGO.getIndex(), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );
            assertNull(
                    wb.getStylesSource().findFont(
                            false, new XSSFColor(IndexedColors.INDIGO, new DefaultIndexedColorMap()), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );

            XSSFFont nf = wb.createFont();
            assertEquals(2, wb.getNumberOfFonts());

            assertEquals(1, nf.getIndex());
            assertEquals(nf, wb.getFontAt(1));

            nf.setBold(false);
            nf.setColor(IndexedColors.INDIGO.getIndex());
            nf.setFontHeight((short) 22);
            nf.setFontName("Thingy");
            nf.setItalic(false);
            nf.setStrikeout(true);
            nf.setTypeOffset((short) 2);
            nf.setUnderline((byte) 2);

            assertEquals(2, wb.getNumberOfFonts());
            assertEquals(nf, wb.getFontAt(1));

            assertNotSame(wb.getFontAt(0), wb.getFontAt(1));

            // Find it now
            assertNotNull(
                    wb.findFont(
                            false, IndexedColors.INDIGO.getIndex(), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );
            assertNotNull(
                    wb.getStylesSource().findFont(
                            false, new XSSFColor(IndexedColors.INDIGO, new DefaultIndexedColorMap()), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );

            XSSFFont font = wb.findFont(
                    false, IndexedColors.INDIGO.getIndex(), (short) 22,
                    "Thingy", false, true, (short) 2, (byte) 2
            );
            assertNotNull(font);
            assertEquals(
                    1,
                    font.getIndex()
            );
            assertEquals(nf,
                    wb.findFont(
                            false, IndexedColors.INDIGO.getIndex(), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );
            assertEquals(nf,
                    wb.getStylesSource().findFont(
                            false, new XSSFColor(IndexedColors.INDIGO, new DefaultIndexedColorMap()), (short) 22,
                            "Thingy", false, true, (short) 2, (byte) 2
                    )
            );

        }
    }

    @Test
    public void testEquals() {
        XSSFFont font = new XSSFFont();
        XSSFFont equ = new XSSFFont();
        XSSFFont notequ = new XSSFFont();
        notequ.setItalic(true);

        assertEquals(equ, font);
        assertNotEquals(font, notequ);

        notequ.setItalic(false);
        notequ.setThemeColor((short) 123);
        assertNotEquals(font, notequ);
    }

    @Test
    public void testBug62272() throws IOException {
        try (XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("62272.xlsx")) {
            // make sure we read the font-color with alpha
            checkFontColor(wb);
        }
    }

    @Test
    public void testBug62272a() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("test");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);

            // create a font with alpha
            XSSFFont font = wb.createFont();
            font.setColor(new XSSFColor(new byte[]{
                    0x7f, 0x33, (byte) 0xCC, 0x66
            }));

            XSSFCellStyle style = wb.createCellStyle();
            style.setFont(font);

            cell.setCellStyle(style);
            cell.setCellValue("testtext");

            // make sure the alpha-value was stored properly
            checkFontColor(wb);

            /*try (OutputStream out = new FileOutputStream("/tmp/testout.xlsx")) {
                wb.write(out);
            }*/
        }
    }

    @Test
    public void testBug62272b() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("test");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);

            // create a font with alpha
            Color androidColor = Color.valueOf(
                    Color.argb(0x7f, 0x33, 0xCC, 0x66));
            XSSFFont font = wb.createFont();
            font.setColor(new XSSFColor(androidColor,
                    null));

            XSSFCellStyle style = wb.createCellStyle();
            style.setFont(font);

            cell.setCellStyle(style);
            cell.setCellValue("testtext");

            // make sure the alpha-value was stored properly
            checkFontColor(wb);

            /*try (OutputStream out = new FileOutputStream("/tmp/testout.xlsx")) {
                wb.write(out);
            }*/
        }
    }

    private static void checkFontColor(Workbook wb) {
        int fontIdx = wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFontIndex();
        Font font = wb.getFontAt(fontIdx);
        //System.out.println(font.getColor());

        CTColor[] colorArray = ((XSSFFont) font).getCTFont().getColorArray();
        //System.out.println(Arrays.toString(colorArray));
        assertArrayEquals(new byte[]{
                0x7f, 0x33, (byte) 0xCC, 0x66
        }, colorArray[0].getRgb());
    }
}