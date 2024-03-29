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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.junit.runner.RunWith;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFOffset {

    @Test
    public void testOffsetWithEmpty23Arguments() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Cell cell = workbook.createSheet().createRow(0).createCell(0);
            cell.setCellFormula("OFFSET(B1,,)");

            String value = "EXPECTED_VALUE";
            Cell valueCell = cell.getRow().createCell(1);
            valueCell.setCellValue(value);

            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            assertEquals(CellType.STRING, cell.getCachedFormulaResultType());
            assertEquals(value, cell.getStringCellValue());
        }
    }
}
