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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf;

import java.io.File;

import org.junit.Test;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.HSSFTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.PackageAccess;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFParser {
    @Test
    public void testXlsx() throws Exception {
        final File file = HSSFTestDataSamples.getSampleFile("github-321.xlsx");
        // unless we use read-only access, the underlying file gets updated
        try (
                OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ);
                XSSFWorkbook wb = XSSFParser.parse(pkg)
        ) {
            assertNotNull(wb);
            assertEquals(1, wb.getNumberOfSheets());
        }
    }

    @Test
    public void testFailOnXls() {
        final File file = HSSFTestDataSamples.getSampleFile("44010-SingleChart.xls");
        XSSFReadException xre = assertThrows(XSSFReadException.class, () -> XSSFParser.parse(file));
        assertTrue(xre.getCause() instanceof OLE2NotOfficeXmlFileException);
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