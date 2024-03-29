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
package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.model;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMapInfo;
import org.w3c.dom.Node;

import java.io.IOException;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ooxml.POIXMLDocumentPart;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFMap;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RunWith(POIJUnit4ClassRunner.class)
public final class TestMapInfo {
    @Test
    public void testMapInfoExists() throws IOException {
        try (XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("CustomXMLMappings.xlsx")) {

            MapInfo mapInfo = null;
            SingleXmlCells singleXMLCells = null;

            for (POIXMLDocumentPart p : wb.getRelations()) {


                if (p instanceof MapInfo) {
                    mapInfo = (MapInfo) p;


                    CTMapInfo ctMapInfo = mapInfo.getCTMapInfo();

                    assertNotNull(ctMapInfo);

                    assertEquals(1, ctMapInfo.sizeOfSchemaArray());

                    for (XSSFMap map : mapInfo.getAllXSSFMaps()) {
                        Node xmlSchema = map.getSchema();
                        assertNotNull(xmlSchema);
                    }
                }
            }

            XSSFSheet sheet1 = wb.getSheetAt(0);

            for (POIXMLDocumentPart p : sheet1.getRelations()) {

                if (p instanceof SingleXmlCells) {
                    singleXMLCells = (SingleXmlCells) p;
                }

            }
            assertNotNull(mapInfo);
            assertNotNull(singleXMLCells);
        }
    }
}
