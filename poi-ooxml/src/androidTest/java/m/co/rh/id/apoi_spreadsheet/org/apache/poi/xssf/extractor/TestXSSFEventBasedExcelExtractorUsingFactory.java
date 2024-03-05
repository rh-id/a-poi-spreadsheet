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
package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.extractor;

import org.junit.After;
import org.junit.runner.RunWith;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.ExtractorFactory;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.HSSFTestDataSamples;

@RunWith(POIJUnit4ClassRunner.class)
public class TestXSSFEventBasedExcelExtractorUsingFactory extends TestXSSFEventBasedExcelExtractor {
    @Override
    protected final XSSFEventBasedExcelExtractor getExtractor(String sampleName) throws Exception {
        ExtractorFactory.setAllThreadsPreferEventExtractors(true);
        return (XSSFEventBasedExcelExtractor) ExtractorFactory.createExtractor(HSSFTestDataSamples.openSampleFileStream(sampleName));
    }

    @After
    public void tearDown() {
        // reset setting to not affect other tests
        ExtractorFactory.setAllThreadsPreferEventExtractors(null);
    }
}
