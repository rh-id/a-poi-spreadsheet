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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j;

import m.co.rh.id.apoi_spreadsheet.base.util.TempFile;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.POIDataSamples;

import java.io.File;
import java.io.InputStream;
import java.io.IOException;

/**
 * Centralises logic for finding/opening sample files for ooxml4j unit tests
 */
public final class OpenXML4JTestDataSamples {
    private static final POIDataSamples _samples = POIDataSamples.getOpenXML4JInstance();

    private OpenXML4JTestDataSamples() {
        // no instances of this class
    }

    public static InputStream openSampleStream(String sampleFileName) {
        return _samples.openResourceAsStream(sampleFileName);
    }
    public static String getSampleFileName(String sampleFileName) {
        return getSampleFile(sampleFileName).getAbsolutePath();
    }

    public static File getSampleFile(String sampleFileName) {
        return _samples.getFile(sampleFileName);
    }

    public static File getOutputFile(String outputFileName) throws IOException {
        int idx = outputFileName.lastIndexOf('.');
        return TempFile.createTempFile(outputFileName.substring(0,idx), outputFileName.substring(idx));
    }
}
