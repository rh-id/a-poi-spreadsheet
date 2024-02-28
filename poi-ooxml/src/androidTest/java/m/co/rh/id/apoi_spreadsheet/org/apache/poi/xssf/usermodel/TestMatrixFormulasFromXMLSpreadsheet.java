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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import android.util.Log;

import org.junit.AfterClass;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4Parameterized;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.ErrorEval;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.functions.BaseTestNumeric;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellValue;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FormulaEvaluator;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LocaleUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.XSSFTestDataSamples;

@RunWith(POIJUnit4Parameterized.class)
public final class TestMatrixFormulasFromXMLSpreadsheet {

    private static final String TAG = "TestMatrixFormulasFromXMLSpreadsheet";

    private static XSSFWorkbook workbook;
    private static Sheet sheet;
    private static FormulaEvaluator evaluator;
    private static Locale userLocale;

    /*
     * Unlike TestFormulaFromSpreadsheet which this class is modified from, there is no
     * differentiation between operators and functions, if more functionality is implemented with
     * array formulas then it might be worth it to separate operators from functions
     *
     * Also, output matrices are statically 3x3, if larger matrices wanted to be tested
     * then adding matrix size parameter would be useful and parsing would be based off that.
     */

    private interface Navigator {
        /**
         * Name of the test spreadsheet (found in the standard test data folder)
         */
        String FILENAME = "MatrixFormulaEvalTestData.xlsx";
        /**
         * Row (zero-based) in the spreadsheet where operations start
         */
        int START_OPERATORS_ROW_INDEX = 1;
        /**
         * Column (zero-based) in the spreadsheet where operations start
         */
        int START_OPERATORS_COL_INDEX = 0;
        /**
         * Column (zero-based) in the spreadsheet where evaluations start
         */
        int START_RESULT_COL_INDEX = 7;
        /**
         * Column separation in the spreadsheet between evaluations and expected results
         */
        int COL_OFF_EXPECTED_RESULT = 3;
        /**
         * Row separation in the spreadsheet between operations
         */
        int ROW_OFF_NEXT_OP = 4;
        /**
         * Used to indicate when there are no more operations left
         */
        String END_OF_TESTS = "<END>";

    }


    @AfterClass
    public static void closeResource() throws Exception {
        LocaleUtil.setUserLocale(userLocale);
        workbook.close();
    }

    @Parameterized.Parameters(name = "{index}: processFunctionRow({0},{1},{2})")
    public static Iterable<Object[]> data() throws InvalidFormatException, IOException {
        // Function "Text" uses custom-formats which are locale specific
        // can't set the locale on a per-testrun execution, as some settings have been
        // already set, when we would try to change the locale by then
        userLocale = LocaleUtil.getUserLocale();
        LocaleUtil.setUserLocale(Locale.ROOT);

        workbook = XSSFTestDataSamples.openSampleWorkbook(Navigator.FILENAME);
        sheet = workbook.getSheetAt(0);
        evaluator = new XSSFFormulaEvaluator(workbook);

        List<Object[]> data = new ArrayList<>();

        processFunctionGroup(data, Navigator.START_OPERATORS_ROW_INDEX, null);

        return data;
    }

    /**
     * @param startRowIndex         row index in the spreadsheet where the first function/operator is found
     * @param testFocusFunctionName name of a single function/operator to test alone.
     *                              Typically pass <code>null</code> to test all functions
     */
    private static void processFunctionGroup(List<Object[]> data, int startRowIndex, String testFocusFunctionName) {
        for (int rowIndex = startRowIndex; true; rowIndex += Navigator.ROW_OFF_NEXT_OP) {
            Row r = sheet.getRow(rowIndex);
            String targetFunctionName = getTargetFunctionName(r);
            assertNotNull("Test spreadsheet cell empty on row ("
                            + (rowIndex) + "). Expected function name or '"
                            + Navigator.END_OF_TESTS + "'",
                    targetFunctionName);
            if (targetFunctionName.equals(Navigator.END_OF_TESTS)) {
                // found end of functions list
                break;
            }
            if (testFocusFunctionName == null || targetFunctionName.equalsIgnoreCase(testFocusFunctionName)) {
                data.add(new Object[]{targetFunctionName, rowIndex});
            }
        }
    }

    @Parameterized.Parameter
    public String targetFunctionName;
    @Parameterized.Parameter(1)
    public int formulasRowIdx;

    @Test
    public void processFunctionRow() {
        int endColNum = Navigator.START_RESULT_COL_INDEX + Navigator.COL_OFF_EXPECTED_RESULT;

        for (int rowNum = formulasRowIdx; rowNum < formulasRowIdx + Navigator.ROW_OFF_NEXT_OP - 1; rowNum++) {
            for (int colNum = Navigator.START_RESULT_COL_INDEX; colNum < endColNum; colNum++) {
                Row r = sheet.getRow(rowNum);

                /* mainly to escape row failures on MDETERM which only returns a scalar */
                if (r == null) {
                    continue;
                }

                Cell c = sheet.getRow(rowNum).getCell(colNum);

                if (c == null || c.getCellType() != CellType.FORMULA) {
                    continue;
                }

                CellValue actValue = evaluator.evaluate(c);
                Cell expValue = sheet.getRow(rowNum).getCell(colNum + Navigator.COL_OFF_EXPECTED_RESULT);

                String msg = String.format(Locale.ROOT, "Function '%s': Formula: %s @ %d:%d"
                        , targetFunctionName, c.getCellFormula(), rowNum, colNum);

                assertNotNull(msg + " - Bad setup data expected value is null", expValue);
                assertNotNull(msg + " - actual value was null", actValue);

                final CellType cellType = expValue.getCellType();
                switch (cellType) {
                    case BLANK:
                        assertEquals(msg, CellType.BLANK, actValue.getCellType());
                        break;
                    case BOOLEAN:
                        assertEquals(msg, CellType.BOOLEAN, actValue.getCellType());
                        assertEquals(msg, expValue.getBooleanCellValue(), actValue.getBooleanValue());
                        break;
                    case ERROR:
                        assertEquals(msg, CellType.ERROR, actValue.getCellType());
                        assertEquals(msg, ErrorEval.getText(expValue.getErrorCellValue()), ErrorEval.getText(actValue.getErrorValue()));
                        break;
                    case FORMULA: // will never be used, since we will call method after formula evaluation
                        fail("Cannot expect formula as result of formula evaluation: " + msg);
                    case NUMERIC:
                        assertEquals(msg, CellType.NUMERIC, actValue.getCellType());
                        BaseTestNumeric.assertDouble(msg, expValue.getNumericCellValue(), actValue.getNumberValue(), BaseTestNumeric.POS_ZERO, BaseTestNumeric.DIFF_TOLERANCE_FACTOR);
                        break;
                    case STRING:
                        assertEquals(msg, CellType.STRING, actValue.getCellType());
                        assertEquals(msg, expValue.getRichStringCellValue().getString(), actValue.getStringValue());
                        break;
                    default:
                        fail("Unexpected cell type: " + cellType);
                }
            }
        }
    }

    /**
     * @return <code>null</code> if cell is missing, empty or blank
     */
    private static String getTargetFunctionName(Row r) {
        if (r == null) {
            Log.w(TAG, "Given null row, can't figure out function name");
            return null;
        }
        Cell cell = r.getCell(Navigator.START_OPERATORS_COL_INDEX);
        Log.d(TAG, String.valueOf(Navigator.START_OPERATORS_COL_INDEX));
        if (cell == null) {
            Log.w(TAG, String.format("Row %d has no cell " + Navigator.START_OPERATORS_COL_INDEX + ", can't figure out function name", r.getRowNum()));
            return null;
        }

        CellType ct = cell.getCellType();
        assertTrue("Bad cell type for 'function name' column: (" + cell.getCellType() + ") row (" + (r.getRowNum() + 1) + ")",
                ct == CellType.BLANK || ct == CellType.STRING);

        return (ct == CellType.STRING) ? cell.getRichStringCellValue().getString() : null;
    }
}
