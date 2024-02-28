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

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertFalse;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.junit.Test;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BorderExtent;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.BorderStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FillPatternType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.PropertyTemplate;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.IndexedColorMap;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFCell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFCellStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFColor;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public class TestSSUtilVsXSSFColor {

    @Test
    public void testXSSFCellStyle() throws Exception {

        try (
                XSSFWorkbook workbook = new XSSFWorkbook();
                UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()
        ) {
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            //final String rgbS = "ffff00";
            final byte[] rgbB = {(byte) 255, (byte) 255, 0};
            IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
            XSSFColor color = new XSSFColor(rgbB, colorMap);
            cellStyle.setFillForegroundColor(color);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            final int startDataRow = 6; // row 7 (index 0-based)
            final int endDataRow = 11; // row 12 (index 0-based)
            final int startDataColumn = 1; // column B (index 0-based)
            final int endDataColumn = 10; // column K (index 0-based)

            Sheet sheet = workbook.createSheet();

            for (int r = startDataRow; r <= endDataRow; r++) {
                Row row = sheet.createRow(r);
                for (int c = startDataColumn; c <= endDataColumn; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue(cell.getAddress().formatAsString());
                    cell.setCellStyle(cellStyle);
                }
            }

            PropertyTemplate propertyTemplate = new PropertyTemplate();
            propertyTemplate.drawBorders(new CellRangeAddress(startDataRow, endDataRow, startDataColumn, endDataColumn),
                    BorderStyle.MEDIUM, BorderExtent.ALL);

            propertyTemplate.applyBorders(sheet); // after this all cell interiors are filled black, because IndexedColors 0 is set
            // same is using all other org.apache.poi.ss.util classes which manipulate cell styles (CellUtil or RegionUtil)

            workbook.write(bos);

            try (XSSFWorkbook wb2 = new XSSFWorkbook(bos.toInputStream())) {
                XSSFSheet sheetWb2 = wb2.getSheetAt(0);
                XSSFCell testCell = sheetWb2.getRow(startDataRow).getCell(startDataColumn);
                XSSFCellStyle testStyle = testCell.getCellStyle();
                XSSFColor testColor = testStyle.getFillForegroundXSSFColor();
                assertFalse(testColor.isIndexed());
                assertArrayEquals(rgbB, testColor.getRGB());
            }
        }
    }
}