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


import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertThrows;

import android.content.Context;

import androidx.test.platform.app.InstrumentationRegistry;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ooxml.POIXMLException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.MemoryLeakVerifier;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.util.TestIOUtils;

/**
 * A test which uses {@link MemoryLeakVerifier} to ensure that certain
 * objects are not left over in memory after the test.
 * <p>
 * E.g. verifies that objects are freed when stuff is removed from sheets or rows
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFMemoryLeak {
    private final MemoryLeakVerifier verifier = new MemoryLeakVerifier();

    // keep some items in memory, so checks in tearDown() actually
    // verify that they do not keep certain objects in memory,
    // e.g. nested CT... objects which should be released
    @SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
    private final List<Object> references = new ArrayList<>();

    @BeforeClass
    public static void setUpClass() {
        POISpreadsheetContext.getInstance().setAppContext(InstrumentationRegistry.getInstrumentation().getTargetContext());
    }

    @AfterClass
    public static void tearDownClass() {
        POISpreadsheetContext.getInstance().setAppContext(null);
        Context context = InstrumentationRegistry.getInstrumentation().getTargetContext();
        TestIOUtils.deleteChildFiles(context.getCacheDir());
    }

    @After
    public void tearDown() {
        verifier.assertGarbageCollected();
    }

    @Test
    public void testWriteRow() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final XSSFSheet sheet1 = wb.createSheet("Sheet1");
            final XSSFRow row = sheet1.createRow(0);
            final XSSFCell cell = row.createCell(0);
            cell.setCellValue("hello");

            // Cannot check the CTCell here as it is reused now and thus
            // not freed until we free up the Cell itself
            //verifier.addObject(ctCell);

            try (OutputStream out = UnsynchronizedByteArrayOutputStream.builder().setBufferSize(8192).get()) {
                wb.write(out);
            }

            CTCell ctCell = cell.getCTCell();
            assertSame("The CTCell should not be replaced", cell.getCTCell(), ctCell);
            assertSame("The CTCell in the row should not be replaced", row.getCTRow().getCArray(0), ctCell);
        }
    }

    @Test
    public void testRemoveCellFromRow() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final XSSFSheet sheet1 = wb.createSheet("Sheet1");
            final XSSFRow rowToCheck = sheet1.createRow(0);
            references.add(rowToCheck);

            XSSFCell cell = rowToCheck.createCell(0);
            cell.setCellValue("hello");

            // previously the CTCell was still referenced in the CTRow, verify that it is freed
            verifier.addObject(cell);
            verifier.addObject(cell.getCTCell());

            rowToCheck.removeCell(cell);

            assertNull(rowToCheck.getCell(0));
        }
    }

    @Test
    public void testRemove2CellsFromRow() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            final XSSFSheet sheet1 = wb.createSheet("Sheet1");
            final XSSFRow rowToCheck = sheet1.createRow(0);
            references.add(rowToCheck);

            XSSFCell cell1 = rowToCheck.createCell(0);
            cell1.setCellValue("hello");
            XSSFCell cell2 = rowToCheck.createCell(1);
            cell2.setCellValue("world");

            // previously the CTCell was still referenced in the CTRow, verify that it is freed
            verifier.addObject(cell1);
            verifier.addObject(cell1.getCTCell());
            verifier.addObject(cell2);
            verifier.addObject(cell2.getCTCell());

            rowToCheck.removeCell(cell2);
            rowToCheck.removeCell(cell1);

            assertNull(rowToCheck.getCell(0));
            assertNull(rowToCheck.getCell(1));
        }
    }

    @Test
    public void testRemoveRowFromSheet() throws IOException {
        try (XSSFWorkbook wb1 = new XSSFWorkbook()) {
            final XSSFSheet sheetToCheck = wb1.createSheet("Sheet1");
            references.add(sheetToCheck);
            final XSSFRow row = sheetToCheck.createRow(0);
            final XSSFCell cell = row.createCell(0);
            cell.setCellValue(1);

            // ensure that the row-data is not kept somewhere in another member
            verifier.addObject(row.getCTRow());
            verifier.addObject(row);
            verifier.addObject(cell.getCTCell());
            verifier.addObject(cell);

            sheetToCheck.removeRow(row);

            assertNull(sheetToCheck.getRow(0));
        }
    }

    @Test
    public void testFileLeak() {
        File file = XSSFTestDataSamples.getSampleFile("xlsx-corrupted.xlsx");
        verifier.addObject(file);
        assertThrows("Should catch exception as the file is corrupted", POIXMLException.class, () -> new XSSFWorkbook(file));
    }

    @Test
    public void testFileLeak2() throws IOException, InvalidFormatException {
        File file = XSSFTestDataSamples.getSampleFile("xlsx-corrupted.xlsx");
        verifier.addObject(file);
        try (OPCPackage pkg = OPCPackage.open(file)) {
            assertThrows("Should catch exception as the file is corrupted", POIXMLException.class, () -> new XSSFWorkbook(pkg));
        }
    }
}