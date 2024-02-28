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
package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import android.util.Log;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BaseTestConditionalFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Color;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ConditionalFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ExtendedColor;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FontFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.PatternFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFITestDataProvider;

/**
 * XSSF-specific Conditional Formatting tests
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFConditionalFormatting extends BaseTestConditionalFormatting {

    public TestXSSFConditionalFormatting() {
        super(XSSFITestDataProvider.instance);
    }

    //https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.databar?view=openxml-2.8.1
    //defaults are 10% and 90%
    @Override
    protected int defaultDataBarMinLength() {
        return 10;
    }

    @Override
    protected int defaultDataBarMaxLength() {
        return 90;
    }

    @Override
    protected void assertColor(String hexExpected, Color actual) {
        assertNotNull("Color must be given", actual);
        XSSFColor color = (XSSFColor) actual;
        if (hexExpected.length() == 8) {
            assertEquals(hexExpected, color.getARGBHex());
        } else {
            assertEquals(hexExpected, color.getARGBHex().substring(2));
        }
    }

    @Test
    public void testRead() throws IOException {
        testRead("WithConditionalFormatting.xlsx");
    }

    @Test
    public void testReadOffice2007() throws IOException {
        testReadOffice2007("NewStyleConditionalFormattings.xlsx");
    }

    private static final android.graphics.Color PEAK_ORANGE = android.graphics.Color.valueOf(android.graphics.Color.rgb(255, 239, 221));

    @Test
    public void testFontFormattingColor() {
        XSSFWorkbook wb = XSSFITestDataProvider.instance.createWorkbook();
        final Sheet sheet = wb.createSheet();

        final SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();

        // the conditional formatting is not automatically added when it is created...
        assertEquals(0, formatting.getNumConditionalFormattings());
        ConditionalFormattingRule formattingRule = formatting.createConditionalFormattingRule("A1");
        assertEquals(0, formatting.getNumConditionalFormattings());

        // adding the formatting makes it available
        int idx = formatting.addConditionalFormatting(new CellRangeAddress[]{}, formattingRule);

        // verify that it can be accessed now
        assertEquals(0, idx);
        assertEquals(1, formatting.getNumConditionalFormattings());
        assertEquals(1, formatting.getConditionalFormattingAt(idx).getNumberOfRules());

        // this is confusing: the rule is not connected to the sheet, changes are not applied
        // so we need to use setRule() explicitly!
        FontFormatting fontFmt = formattingRule.createFontFormatting();
        assertNotNull(formattingRule.getFontFormatting());
        assertEquals(1, formatting.getConditionalFormattingAt(idx).getNumberOfRules());
        formatting.getConditionalFormattingAt(idx).setRule(0, formattingRule);
        assertNotNull(formatting.getConditionalFormattingAt(idx).getRule(0).getFontFormatting());

        fontFmt.setFontStyle(true, false);

        assertEquals(-1, fontFmt.getFontColorIndex());

        //fontFmt.setFontColorIndex((short)11);
        final ExtendedColor extendedColor = new XSSFColor(PEAK_ORANGE, wb.getStylesSource().getIndexedColors());
        fontFmt.setFontColor(extendedColor);

        PatternFormatting patternFmt = formattingRule.createPatternFormatting();
        assertNotNull(patternFmt);
        patternFmt.setFillBackgroundColor(extendedColor);

        assertEquals(1, formatting.getConditionalFormattingAt(0).getNumberOfRules());
        assertNotNull(formatting.getConditionalFormattingAt(0).getRule(0).getFontFormatting());
        assertNotNull(formatting.getConditionalFormattingAt(0).getRule(0).getFontFormatting().getFontColor());
        assertNotNull(formatting.getConditionalFormattingAt(0).getRule(0).getPatternFormatting().getFillBackgroundColorColor());

        checkFontFormattingColorWriteOutAndReadBack(wb, extendedColor);
    }

    private void checkFontFormattingColorWriteOutAndReadBack(Workbook wb, ExtendedColor extendedColor) {
        Workbook wbBack = XSSFITestDataProvider.instance.writeOutAndReadBack(wb);
        assertNotNull(wbBack);

        assertEquals(1, wbBack.getSheetAt(0).getSheetConditionalFormatting().getNumConditionalFormattings());
        final ConditionalFormatting formattingBack = wbBack.getSheetAt(0).getSheetConditionalFormatting().getConditionalFormattingAt(0);
        assertEquals(1, wbBack.getSheetAt(0).getSheetConditionalFormatting().getConditionalFormattingAt(0).getNumberOfRules());
        final ConditionalFormattingRule ruleBack = formattingBack.getRule(0);
        final FontFormatting fontFormattingBack = ruleBack.getFontFormatting();
        assertNotNull(formattingBack);
        assertNotNull(fontFormattingBack.getFontColor());
        assertEquals(extendedColor, fontFormattingBack.getFontColor());
        assertEquals(extendedColor, ruleBack.getPatternFormatting().getFillBackgroundColorColor());
    }

    @Override
    protected boolean applyLimitOf3() {
        return false;
    }

    @Override
    protected void checkThreshold(ConditionalFormattingThreshold threshold) {
        assertNull(threshold.getValue());
        assertNull(threshold.getFormula());
        assertTrue("threshold is a XSSFConditionalFormattingThreshold?",
                threshold instanceof XSSFConditionalFormattingThreshold);
        XSSFConditionalFormattingThreshold xssfThreshold = (XSSFConditionalFormattingThreshold) threshold;
        assertTrue("gte defaults to true?",
                xssfThreshold.isGte());
        xssfThreshold.setGte(false);
        assertFalse("gte changed to false?",
                xssfThreshold.isGte());
        xssfThreshold.setGte(true);
        assertTrue("gte changed to true?",
                xssfThreshold.isGte());
    }
}
