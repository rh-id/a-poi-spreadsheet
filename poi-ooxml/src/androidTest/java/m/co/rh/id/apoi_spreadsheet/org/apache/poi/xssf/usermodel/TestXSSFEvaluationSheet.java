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
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.IOException;
import java.util.AbstractMap;
import java.util.Map;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.EvaluationSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BaseTestXEvaluationSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFEvaluationSheet extends BaseTestXEvaluationSheet {

    @Test
    public void testSheetEval() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("test");
        XSSFRow row = sheet.createRow(0);
        row.createCell(0);
        XSSFEvaluationSheet evalsheet = new XSSFEvaluationSheet(sheet);

        assertNotNull("Cell 0,0 is found", evalsheet.getCell(0, 0));
        assertNull("Cell 0,1 is not found", evalsheet.getCell(0, 1));
        assertNull("Cell 1,0 is not found", evalsheet.getCell(1, 0));

        // now add Cell 0,1
        row.createCell(1);

        assertNotNull("Cell 0,0 is found", evalsheet.getCell(0, 0));
        assertNotNull("Cell 0,1 is now also found", evalsheet.getCell(0, 1));
        assertNull("Cell 1,0 is not found", evalsheet.getCell(1, 0));

        // after clearing all values it also works
        row.createCell(2);
        evalsheet.clearAllCachedResultValues();

        assertNotNull("Cell 0,0 is found", evalsheet.getCell(0, 0));
        assertNotNull("Cell 0,2 is now also found", evalsheet.getCell(0, 2));
        assertNull("Cell 1,0 is not found", evalsheet.getCell(1, 0));

        // other things
        assertEquals(sheet, evalsheet.getXSSFSheet());
    }

    @Test
    public void testBug65675() throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setIgnoreMissingWorkbooks(true);

            XSSFSheet sheet = workbook.createSheet("sheet");
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell = row.createCell(0, CellType.FORMULA);

            try {
                cell.setCellFormula("[some-workbook-that-does-not-yet-exist.xlsx]main!B:D");
                //it might be better if this succeeded but just adding this regression test for now
                fail("expected exception");
            } catch (RuntimeException re) {
                assertEquals("Book not linked for filename some-workbook-that-does-not-yet-exist.xlsx", re.getMessage());
            }
        }
    }

    @Override
    protected Map.Entry<Sheet, EvaluationSheet> getInstance() {
        XSSFSheet sheet = new XSSFWorkbook().createSheet();
        return new AbstractMap.SimpleEntry<>(sheet, new XSSFEvaluationSheet(sheet));
    }
}
