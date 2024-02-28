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

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.HSSFTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.util.ZipSecureFile;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFFileChecks {
    @Test
    public void testWithReducedFileLimit() {
        final long defaultLimit = ZipSecureFile.getMaxFileCount();
        ZipSecureFile.setMaxFileCount(5);
        try (InputStream is = HSSFTestDataSamples.openSampleFileStream("HeaderFooterComplexFormats.xlsx")) {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            fail("expected IOException");
        } catch (IOException e) {
            assertTrue("unexpected exception message: " + e.getMessage(),
                    e.getMessage().contains("ZipSecureFile.setMaxFileCount()"));
        } finally {
            ZipSecureFile.setMaxFileCount(defaultLimit);
        }
    }

    @Test
    public void testFileWithReducedFileLimit() throws IOException {
        final File file = HSSFTestDataSamples.getSampleFile("HeaderFooterComplexFormats.xlsx");
        final long defaultLimit = ZipSecureFile.getMaxFileCount();
        ZipSecureFile.setMaxFileCount(5);
        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file);
            fail("expected InvalidFormatException");
        } catch (InvalidFormatException e) {
            assertTrue("unexpected exception message: " + e.getMessage(),
                    e.getMessage().contains("ZipSecureFile.setMaxFileCount()"));
        } finally {
            ZipSecureFile.setMaxFileCount(defaultLimit);
        }
    }

    @Test
    public void testWithGraceEntrySize() throws IOException {
        final double defaultInflateRatio = ZipSecureFile.getMinInflateRatio();
        // setting MinInflateRatio but the default GraceEntrySize will mean this is ignored
        // this exception will not happen with the default GraceEntrySize
        ZipSecureFile.setMinInflateRatio(0.50);
        try (InputStream is = HSSFTestDataSamples.openSampleFileStream("HeaderFooterComplexFormats.xlsx")) {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            assertNotNull(xssfWorkbook);
        } finally {
            ZipSecureFile.setMinInflateRatio(defaultInflateRatio);
        }
    }

    @Test
    public void testWithReducedGraceEntrySize() {
        final long defaultGraceSize = ZipSecureFile.getGraceEntrySize();
        final double defaultInflateRatio = ZipSecureFile.getMinInflateRatio();
        ZipSecureFile.setGraceEntrySize(0);
        // setting MinInflateRatio to cause an exception
        // this exception will not happen with the default GraceEntrySize
        ZipSecureFile.setMinInflateRatio(0.50);
        try (InputStream is = HSSFTestDataSamples.openSampleFileStream("HeaderFooterComplexFormats.xlsx")) {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            fail("expected IOException");
        } catch (IOException e) {
            assertTrue("unexpected exception message: " + e.getMessage(),
                    e.getMessage().contains("ZipSecureFile.setMinInflateRatio()"));
        } finally {
            ZipSecureFile.setMinInflateRatio(defaultInflateRatio);
            ZipSecureFile.setGraceEntrySize(defaultGraceSize);
        }
    }

}
