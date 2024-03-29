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

import org.junit.Test;

import java.util.Map;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.EvaluationSheet;

public abstract class BaseTestXEvaluationSheet {
    /**
     * Get a pair of underlying sheet and evaluation sheet.
     */
    protected abstract Map.Entry<Sheet, EvaluationSheet> getInstance();

    @Test
    public void lastRowNumIsUpdatedFromUnderlyingSheet_bug62993() {
        Map.Entry<Sheet, EvaluationSheet> sheetPair = getInstance();
        Sheet underlyingSheet = sheetPair.getKey();
        EvaluationSheet instance = sheetPair.getValue();

        assertEquals(-1, instance.getLastRowNum());

        underlyingSheet.createRow(0);
        underlyingSheet.createRow(1);
        underlyingSheet.createRow(2);
        assertEquals(2, instance.getLastRowNum());

        underlyingSheet.removeRow(underlyingSheet.getRow(2));
        assertEquals(1, instance.getLastRowNum());
    }

}
