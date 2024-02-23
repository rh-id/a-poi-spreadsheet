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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.EvaluationWorkbook.ExternalSheet;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.ptg.NamePtg;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.ptg.NameXPtg;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;

/**
 * Abstracts a workbook for the purpose of converting formula to text.
 */
@Internal
public interface FormulaRenderingWorkbook {
    /**
     * @return {@code null} if externSheetIndex refers to a sheet inside the current workbook
     */
    ExternalSheet getExternalSheet(int externSheetIndex);

    /**
     * @return the name of the (first) sheet referred to by the given external sheet index
     */
    String getSheetFirstNameByExternSheet(int externSheetIndex);
    /**
     * @return the name of the (last) sheet referred to by the given external sheet index
     */
    String getSheetLastNameByExternSheet(int externSheetIndex);

    String resolveNameXText(NameXPtg nameXPtg);
    String getNameText(NamePtg namePtg);
}
