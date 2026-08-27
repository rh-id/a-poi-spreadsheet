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
package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util;

import org.junit.Test;
import org.junit.runner.RunWith;

import java.util.Arrays;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellPropertyType;

/**
 * Test for CellUtil constants
 */
@RunWith(POIJUnit4ClassRunner.class)
public class TestCellUtil {
    @Test
    public void testNamePropertyMap() {
        Arrays.stream(CellPropertyType.values()).forEach(cellPropertyType ->
                assertTrue(CellUtil.namePropertyMap.containsValue(cellPropertyType),
                        "missing " + cellPropertyType));
    }

    private void assertTrue(boolean condition, String message) {
        org.junit.Assert.assertTrue(message, condition);
    }
}