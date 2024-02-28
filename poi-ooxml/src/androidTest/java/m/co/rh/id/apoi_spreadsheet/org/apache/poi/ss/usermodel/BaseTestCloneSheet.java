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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThrows;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.ITestDataProvider;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;

/**
 * Common superclass for testing implementations of
 * Workbook.cloneSheet()
 */
public abstract class BaseTestCloneSheet {

    private final ITestDataProvider _testDataProvider;

    protected BaseTestCloneSheet(ITestDataProvider testDataProvider) {
        _testDataProvider = testDataProvider;
    }

    @Test
    public void testCloneSheetBasic() throws IOException {
        Workbook b = _testDataProvider.createWorkbook();
        Sheet s = b.createSheet("Test");
        assertEquals(0, s.addMergedRegion(new CellRangeAddress(0, 1, 0, 1)));
        Sheet clonedSheet = b.cloneSheet(0);

        assertEquals("One merged area", 1, clonedSheet.getNumMergedRegions());

        b.close();
    }

    /**
     * Ensures that pagebreak cloning works properly
     */
    @Test
    public void testPageBreakClones() throws IOException {
        Workbook b = _testDataProvider.createWorkbook();
        Sheet s = b.createSheet("Test");
        s.setRowBreak(3);
        s.setColumnBreak((short) 6);

        Sheet clone = b.cloneSheet(0);
        assertTrue("Row 3 not broken", clone.isRowBroken(3));
        assertTrue("Column 6 not broken", clone.isColumnBroken((short) 6));

        s.removeRowBreak(3);

        assertTrue("Row 3 still should be broken", clone.isRowBroken(3));

        b.close();
    }

    @Test
    public void testCloneSheetIntValid() {
        Workbook wb = _testDataProvider.createWorkbook();
        wb.createSheet("Sheet01");
        wb.cloneSheet(0);
        assertEquals(2, wb.getNumberOfSheets());
        assertThrows(IllegalArgumentException.class, () -> wb.cloneSheet(2));
    }

    @Test
    public void testCloneSheetIntInvalid() {
        Workbook wb = _testDataProvider.createWorkbook();
        wb.createSheet("Sheet01");
        assertThrows(IllegalArgumentException.class, () -> wb.cloneSheet(1));
        assertEquals(1, wb.getNumberOfSheets());
    }
}
