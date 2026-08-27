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
 *  ====================================================================
 */
// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Test creates cells in reverse column order in XSSF and SXSSF and expects
 * saved files to have fixed the order.
 *
 * This is necessary because if columns in the saved file are out of order
 * Excel will show a repair dialog when opening the file and removing data.
 */
@RunWith(POIJUnit4ClassRunner.class)
public final class TestOutOfOrderColumns {

    @Test
    public void outOfOrderColumnsXSSF() throws IOException {
        try (
                XSSFWorkbook wb = new XSSFWorkbook();
                ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ) {
            XSSFSheet sheet = wb.createSheet();

            Row row = sheet.createRow(0);
            // create cells in reverse order
            row.createCell(1).setCellValue("def");
            row.createCell(0).setCellValue("abc");

            wb.write(bos);

            validateOrder(new ByteArrayInputStream(bos.toByteArray()));
        }
    }

    @Test
    public void outOfOrderColumnsSXSSF() throws IOException {
        try (
                SXSSFWorkbook wb = new SXSSFWorkbook();
                ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ) {
            Sheet sheet = wb.createSheet();

            Row row = sheet.createRow(0);
            // create cells in reverse order
            row.createCell(1).setCellValue("xyz");
            row.createCell(0).setCellValue("uvw");

            wb.write(bos);

            validateOrder(new ByteArrayInputStream(bos.toByteArray()));
        }
    }

    @Test
    /** this is the problematic case, as XSSF column sorting is skipped when saving with SXSSF. */
    public void mixOfXSSFAndSXSSF() throws IOException {
        try (
                XSSFWorkbook wb = new XSSFWorkbook();
                ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ) {
            XSSFSheet sheet = wb.createSheet();

            Row row1 = sheet.createRow(0);
            // create cells in reverse order
            row1.createCell(1).setCellValue("def");
            row1.createCell(0).setCellValue("abc");

            try (SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(wb)) {
                Sheet sSheet = sxssfWorkbook.getSheetAt(0);
                Row row2 = sSheet.createRow(1);
                // create cells in reverse order
                row2.createCell(1).setCellValue("xyz");
                row2.createCell(0).setCellValue("uvw");

                sxssfWorkbook.write(bos);

                validateOrder(new ByteArrayInputStream(bos.toByteArray()));
            }
        }
    }

    private void validateOrder(InputStream is) throws IOException {
        // test if saved cells are in order
        try (XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is)) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

            Row resultRow = xssfSheet.getRow(0);
            // POI doesn't show stored order because _cells TreeMap sorts it automatically.
            // The only way to test is to compare the xml.
            String s = resultRow.toString();
            assertTrue(s.matches("(?s).*A1.*B1.*"), "unexpected order: " + s);
        }
    }

    private void assertTrue(boolean condition, String message) {
        org.junit.Assert.assertTrue(message, condition);
    }

}