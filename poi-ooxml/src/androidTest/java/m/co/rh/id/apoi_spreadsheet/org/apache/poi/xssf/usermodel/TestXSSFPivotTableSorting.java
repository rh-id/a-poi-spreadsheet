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

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DataConsolidateFunction;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.AreaReference;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellReference;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFITestDataProvider;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotAreaReference;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFieldSortType;

import java.io.IOException;

import static org.junit.Assert.assertEquals;

/**
 * Test pivot tables created by area reference
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFPivotTableSorting {
    private static final XSSFITestDataProvider _testDataProvider = XSSFITestDataProvider.instance;

    @Test
    public void testNestedSorting() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet();

            Row row0 = sheet.createRow(0);
            // Create a cell and put a value in it.
            row0.createCell(0).setCellValue("Month");
            row0.createCell(1).setCellValue("Name");
            row0.createCell(2).setCellValue("Product");
            row0.createCell(3).setCellValue("Amount");

            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("Jan");
            row1.createCell(1).setCellValue("John");
            row1.createCell(2).setCellValue("Pen");
            row1.createCell(3).setCellValue(5);

            Row row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("Jan");
            row2.createCell(1).setCellValue("Mary");
            row2.createCell(2).setCellValue("Paper");
            row2.createCell(3).setCellValue(5);

            Row row3 = sheet.createRow(3);
            row3.createCell(0).setCellValue("Feb");
            row3.createCell(1).setCellValue("John");
            row3.createCell(2).setCellValue("Clips");
            row3.createCell(3).setCellValue(5);

            Row row4 = sheet.createRow(4);
            row4.createCell(0).setCellValue("Feb");
            row4.createCell(1).setCellValue("Mary");
            row4.createCell(2).setCellValue("Book");
            row4.createCell(3).setCellValue(15);


            AreaReference source = wb.getCreationHelper().createAreaReference("A1:D5");
            XSSFPivotTable pivotTable = sheet.createPivotTable(source, new CellReference("H1"));

            int monthCol = 0;
            int nameCol = 1;
            int productCol = 2;
            int amountCol = 3;

            // Names
            pivotTable.addRowLabel(nameCol);
            pivotTable.addRowLabel(productCol);
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, amountCol);

            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(nameCol).setSortType(STFieldSortType.ASCENDING);
            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(productCol).setSortType(STFieldSortType.DESCENDING);

            // add sorting ASC by sum
            int advancedSortingColumnInPivot = 1;
            CTPivotAreaReference reference =
                    pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(advancedSortingColumnInPivot)
                            .addNewAutoSortScope().addNewPivotArea().addNewReferences().addNewReference();

            // I have no idea what these constants are for
            reference.setField(4294967294L); // if you are curious it's a 2^32 - 2 or just signed -2
            reference.addNewX().setV(0);
            assertEquals(2, pivotTable.getRowLabelColumns().size());
        }
    }
}
