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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel;

import static org.junit.Assert.assertNotNull;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFRow;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellCopyContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellCopyPolicy;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CreationHelper;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DateUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;

@RunWith(POIJUnit4ClassRunner.class)
public class TestHSSFRowCopyRowFrom {
    @Test
    public void testCopyFrom() throws IOException {
        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy.Builder()
                .cellFormula(false) // NOTE: setting to false allows for copying the evaluated formula value.
                .cellStyle(CellCopyPolicy.DEFAULT_COPY_CELL_STYLE_POLICY)
                .cellValue(CellCopyPolicy.DEFAULT_COPY_CELL_VALUE_POLICY)
                .condenseRows(CellCopyPolicy.DEFAULT_CONDENSE_ROWS_POLICY)
                .copyHyperlink(CellCopyPolicy.DEFAULT_COPY_HYPERLINK_POLICY)
                .mergeHyperlink(CellCopyPolicy.DEFAULT_MERGE_HYPERLINK_POLICY)
                .mergedRegions(CellCopyPolicy.DEFAULT_COPY_MERGED_REGIONS_POLICY)
                .rowHeight(CellCopyPolicy.DEFAULT_COPY_ROW_HEIGHT_POLICY)
                .build();

        final LocalDateTime localDateTime = LocalDateTime.of(2023, 1, 1, 0, 0, 0);
        final LocalDateTime nonValidExcelDate = LocalDateTime.of(1899, 12, 31, 0, 0, 0);
        final Object[][] data = {
                {"transaction_id", "transaction_date", "transaction_time"},
                {75, localDateTime, nonValidExcelDate.plusHours(9).plusMinutes(53).plusSeconds(44).toLocalTime()},
                {78, localDateTime, nonValidExcelDate.plusHours(9).plusMinutes(55).plusSeconds(16).toLocalTime()}
        };

        final ByteArrayOutputStream workbookOutputStream = new ByteArrayOutputStream();
        try (Workbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("SomeSheetName");
            populateSheet(sheet, data);
            setCellStyles(sheet, workbook);
            workbook.write(workbookOutputStream);
        }

        try (HSSFWorkbook originalWorkbook = new HSSFWorkbook(new ByteArrayInputStream(workbookOutputStream.toByteArray()))) {
            final Iterator<Sheet> originalSheetsIterator = originalWorkbook.sheetIterator();
            final CellCopyContext cellCopyContext = new CellCopyContext();

            while (originalSheetsIterator.hasNext()) {
                final HSSFSheet originalSheet = (HSSFSheet) originalSheetsIterator.next();
                final String originalSheetName = originalSheet.getSheetName();
                final Iterator<Row> originalRowsIterator = originalSheet.rowIterator();

                try (HSSFWorkbook newWorkbook = new HSSFWorkbook()) {
                    final HSSFSheet newSheet = newWorkbook.createSheet(originalSheetName);
                    while (originalRowsIterator.hasNext()) {
                        HSSFRow originalRow = (HSSFRow) originalRowsIterator.next();
                        HSSFRow newRow = newSheet.createRow(originalRow.getRowNum());
                        newRow.copyRowFrom(originalRow, cellCopyPolicy, cellCopyContext);
                    }
                }
            }
        }
    }

    private static void populateSheet(Sheet sheet, Object[][] data) {
        int rowCount = 0;
        for (Object[] dataRow : data) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;

            for (Object field : dataRow) {
                Cell cell = row.createCell(columnCount++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                } else if (field instanceof Long) {
                    cell.setCellValue((Long) field);
                } else if (field instanceof LocalDateTime) {
                    cell.setCellValue((LocalDateTime) field);
                } else if (field instanceof LocalTime) {
                    cell.setCellValue(DateUtil.convertTime(DateTimeFormatter.ISO_LOCAL_TIME.format((LocalTime) field)));
                }
            }
        }
    }

    private void setCellStyles(Sheet sheet, Workbook workbook) {
        CreationHelper creationHelper = workbook.getCreationHelper();
        CellStyle dayMonthYearCellStyle = workbook.createCellStyle();
        dayMonthYearCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/mm/yyyy"));
        CellStyle hourMinuteSecond = workbook.createCellStyle();
        hourMinuteSecond.setDataFormat((short) 21); // 21 represents format h:mm:ss
        for (int rowNum = sheet.getFirstRowNum() + 1; rowNum < sheet.getLastRowNum() + 1; rowNum++) {
            Row row = sheet.getRow(rowNum);
            assertNotNull(row);
            assertNotNull(row.getCell(1));
            assertNotNull(row.getCell(2));
            row.getCell(1).setCellStyle(dayMonthYearCellStyle);
            row.getCell(2).setCellStyle(hourMinuteSecond);
        }
    }
}