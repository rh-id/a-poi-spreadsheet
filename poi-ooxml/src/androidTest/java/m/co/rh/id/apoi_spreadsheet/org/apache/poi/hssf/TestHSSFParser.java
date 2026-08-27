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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf;

import java.io.InputStream;

import org.junit.Test;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.filesystem.OfficeXmlFileException;

@RunWith(POIJUnit4ClassRunner.class)
public class TestHSSFParser {
    @Test
    public void testXls() throws Exception {
        try (
                InputStream stream = HSSFTestDataSamples.openSampleFileStream("44010-SingleChart.xls");
                HSSFWorkbook wb = HSSFParser.parse(stream)
        ) {
            assertNotNull(wb);
            assertEquals(2, wb.getNumberOfSheets());
        }
    }

    @Test
    public void testFailOnXlsx() throws Exception {
        try (InputStream stream = HSSFTestDataSamples.openSampleFileStream("github-321.xlsx")) {
            HSSFReadException hre = assertThrows(HSSFReadException.class, () -> HSSFParser.parse(stream));
            assertTrue(hre.getCause() instanceof OfficeXmlFileException);
        }
    }

    private void assertNotNull(Object obj) {
        org.junit.Assert.assertNotNull(obj);
    }

    private void assertEquals(int expected, int actual) {
        org.junit.Assert.assertEquals(expected, actual);
    }

    private <T extends Throwable> T assertThrows(Class<T> expectedType, org.junit.function.ThrowingRunnable runnable) {
        return org.junit.Assert.assertThrows(expectedType, runnable);
    }

    private void assertTrue(boolean condition) {
        org.junit.Assert.assertTrue(condition);
    }
}