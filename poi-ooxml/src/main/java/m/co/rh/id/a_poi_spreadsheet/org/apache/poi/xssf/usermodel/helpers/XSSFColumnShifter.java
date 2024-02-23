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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.usermodel.helpers;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaShifter;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.helpers.ColumnShifter;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Helper for shifting columns up or down
 *
 * @since POI 4.0.0
 */
// non-Javadoc: When possible, code should be implemented in the ColumnShifter abstract class to avoid duplication with
// {@link org.apache.poi.hssf.usermodel.helpers.HSSFColumnShifter}
@Beta
public final class XSSFColumnShifter extends ColumnShifter {

    public XSSFColumnShifter(XSSFSheet sh) {
        super(sh);
    }

    @Override
    public void updateNamedRanges(FormulaShifter formulaShifter) {
        XSSFRowColShifter.updateNamedRanges(sheet, formulaShifter);
    }

    @Override
    public void updateFormulas(FormulaShifter formulaShifter) {
        XSSFRowColShifter.updateFormulas(sheet, formulaShifter);
    }

    @Override
    public void updateConditionalFormatting(FormulaShifter formulaShifter) {
        XSSFRowColShifter.updateConditionalFormatting(sheet, formulaShifter);
    }

    @Override
    public void updateHyperlinks(FormulaShifter formulaShifter) {
        XSSFRowColShifter.updateHyperlinks(sheet, formulaShifter);
    }
}