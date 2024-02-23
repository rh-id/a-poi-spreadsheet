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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.helpers;

import android.util.Log;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFCell;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFRow;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFSheet;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaParseException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaParser;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaRenderer;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaShifter;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaType;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.ptg.Ptg;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;

/**
 * Class for code common to {@link HSSFRowShifter} and {@link HSSFColumnShifter}
 *
 * @since POI 5.1.0
 */
@Internal
/*private*/ final class HSSFRowColShifter {
    private static final String TAG = "HSSFRowColShifter";

    private HSSFRowColShifter() { /*no instances for static classes*/}

    /**
     * Update formulas.
     */
    /*package*/
    static void updateFormulas(Sheet sheet, FormulaShifter formulaShifter) {
        //update formulas on the parent sheet
        updateSheetFormulas(sheet, formulaShifter);

        //update formulas on other sheets
        Workbook wb = sheet.getWorkbook();
        for (Sheet sh : wb) {
            if (sheet == sh) continue;
            updateSheetFormulas(sh, formulaShifter);
        }
    }

    /*package*/
    static void updateSheetFormulas(Sheet sh, FormulaShifter formulashifter) {
        for (Row r : sh) {
            HSSFRow row = (HSSFRow) r;
            updateRowFormulas(row, formulashifter);
        }
    }

    /**
     * Update the formulas in specified row using the formula shifting policy specified by shifter
     *
     * @param row            the row to update the formulas on
     * @param formulaShifter the formula shifting policy
     */
    /*package*/
    static void updateRowFormulas(HSSFRow row, FormulaShifter formulaShifter) {
        HSSFSheet sheet = row.getSheet();
        for (Cell c : row) {
            HSSFCell cell = (HSSFCell) c;
            String formula = cell.getCellFormula();
            if (formula.length() > 0) {
                String shiftedFormula = shiftFormula(row, formula, formulaShifter);
                cell.setCellFormula(shiftedFormula);
            }
        }
    }

    /**
     * Shift a formula using the supplied FormulaShifter
     *
     * @param row            the row of the cell this formula belongs to. Used to get a reference to the parent workbook.
     * @param formula        the formula to shift
     * @param formulaShifter the FormulaShifter object that operates on the parsed formula tokens
     * @return the shifted formula if the formula was changed,
     * <code>null</code> if the formula wasn't modified
     */
    /*package*/
    static String shiftFormula(Row row, String formula, FormulaShifter formulaShifter) {
        Sheet sheet = row.getSheet();
        Workbook wb = sheet.getWorkbook();
        int sheetIndex = wb.getSheetIndex(sheet);
        final int rowIndex = row.getRowNum();
        HSSFEvaluationWorkbook fpb = HSSFEvaluationWorkbook.create((HSSFWorkbook) wb);

        try {
            Ptg[] ptgs = FormulaParser.parse(formula, fpb, FormulaType.CELL, sheetIndex, rowIndex);
            String shiftedFmla;
            if (formulaShifter.adjustFormula(ptgs, sheetIndex)) {
                shiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
            } else {
                shiftedFmla = formula;
            }
            return shiftedFmla;
        } catch (FormulaParseException fpe) {
            // Log, but don't change, rather than breaking
            Log.w(TAG, String.format("Error shifting formula on row %d", row.getRowNum()), fpe);
            return formula;
        }
    }

}
