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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.io;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.util.List;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.POIDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFPictureData;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public class TestLoadSaveXSSF {
    private static final POIDataSamples _ssSamples = POIDataSamples.getSpreadSheetInstance();

    @Test
    public void testLoadSample() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(_ssSamples.openResourceAsStream("sample.xlsx"))) {
            assertEquals(3, workbook.getNumberOfSheets());
            assertEquals("Sheet1", workbook.getSheetName(0));
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell = row.getCell((short) 1);
            assertNotNull(cell);
            assertEquals(111.0, cell.getNumericCellValue(), 0.0);
            cell = row.getCell((short) 0);
            assertEquals("Lorem", cell.getRichStringCellValue().getString());
        }
    }

    @Test
    public void testLoadStyles() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(_ssSamples.openResourceAsStream("styles.xlsx"))) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell = row.getCell((short) 0);
            CellStyle style = cell.getCellStyle();
            assertNotNull(style);
        }
    }

    @Test
    public void testLoadPictures() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(_ssSamples.openResourceAsStream("picture.xlsx"))) {
            List<XSSFPictureData> pictures = workbook.getAllPictures();
            assertEquals(1, pictures.size());
        }
    }
}