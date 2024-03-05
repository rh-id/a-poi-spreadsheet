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

import org.junit.After;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.tests.usermodel.BaseTestXRow;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.SXSSFITestDataProvider;

/**
 * Tests for XSSFRow
 */
@RunWith(POIJUnit4ClassRunner.class)
public final class TestSXSSFRow extends BaseTestXRow {

    public TestSXSSFRow() {
        super(SXSSFITestDataProvider.instance);
    }


    @After
    public void tearDown() {
        ((SXSSFITestDataProvider) _testDataProvider).cleanup();
    }

    @Override
    @Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030#c1>")
    public void testCellShiftingRight() {
        // Remove when SXSSFRow.shiftCellsRight() is implemented.
    }

    @Override
    @Ignore("see <https://bz.apache.org/bugzilla/show_bug.cgi?id=62030#c1>")
    public void testCellShiftingLeft() {
        // Remove when SXSSFRow.shiftCellsLeft() is implemented.
    }

    @Test
    public void testCellColumn() throws IOException {
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            SXSSFSheet sheet = wb.createSheet();
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell5 = row.createCell(5);
            assertEquals(5, cell5.getColumnIndex());
        }
    }

}
