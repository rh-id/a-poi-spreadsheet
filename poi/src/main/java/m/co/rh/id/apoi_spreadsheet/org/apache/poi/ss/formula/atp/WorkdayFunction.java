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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.atp;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.OperationEvaluationContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.ErrorEval;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.EvaluationException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.NumberEval;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.ValueEval;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.functions.FreeRefFunction;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DateUtil;

/**
 * Implementation of Excel 'Analysis ToolPak' function WORKDAY()<br>
 * Returns the date past a number of workdays beginning at a start date, considering an interval of holidays. A workday is any non
 * saturday/sunday date.
 * <p>
 * <b>Syntax</b><br>
 * <b>WORKDAY</b>(<b>startDate</b>, <b>days</b>, holidays)
 * <p>
 */
final class WorkdayFunction implements FreeRefFunction {

    public static final FreeRefFunction instance = new WorkdayFunction(ArgumentsEvaluator.instance);

    private ArgumentsEvaluator evaluator;

    private WorkdayFunction(ArgumentsEvaluator anEvaluator) {
        // enforces singleton
        this.evaluator = anEvaluator;
    }

    /**
     * Evaluate for WORKDAY. Given a date, a number of days and an optional date or interval of holidays, determines which date it is past
     * number of parameterized workdays.
     *
     * @return {@link ValueEval} with date as its value.
     */
    public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
        if (args.length < 2 || args.length > 3) {
            return ErrorEval.VALUE_INVALID;
        }

        int srcCellRow = ec.getRowIndex();
        int srcCellCol = ec.getColumnIndex();

        double start;
        int days;
        double[] holidays;
        try {
            start = this.evaluator.evaluateDateArg(args[0], srcCellRow, srcCellCol);
            days = (int) Math.floor(this.evaluator.evaluateNumberArg(args[1], srcCellRow, srcCellCol));
            ValueEval holidaysCell = args.length == 3 ? args[2] : null;
            holidays = this.evaluator.evaluateDatesArg(holidaysCell, srcCellRow, srcCellCol);
            return new NumberEval(DateUtil.getExcelDate(WorkdayCalculator.instance.calculateWorkdays(start, days, holidays)));
        } catch (EvaluationException e) {
            return ErrorEval.VALUE_INVALID;
        }
    }

}
