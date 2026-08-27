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
// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.functions;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;
import java.time.LocalDate;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DateUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.Utils;

/**
 * Tests WORKDAY(start_date, days, [holidays]) excel function
 * https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestWorkdayFunc {
    @Test
    public void testCellRefs() throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row row0 = sheet.createRow(0);
            Cell cellA1 = row0.createCell(0);
            Cell cellB1 = row0.createCell(1);
            cellA1.setCellValue(LocalDate.parse("2024-10-29"));
            cellB1.setCellValue(5);
            Cell cellResult = sheet.createRow(1).createCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            Utils.assertDouble(fe, cellResult, "WORKDAY(A1,B1)", DateUtil.getExcelDate(LocalDate.parse("2024-11-05")));
        }
    }
}