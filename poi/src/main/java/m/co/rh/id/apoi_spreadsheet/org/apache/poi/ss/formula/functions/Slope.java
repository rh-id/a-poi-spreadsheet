/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.functions;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.eval.ValueEval;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.functions.LinearRegressionFunction.FUNCTION;

/**
 * Implementation of Excel function SLOPE()<p>
 *
 * Calculates the SLOPE of the linear regression line that is used to predict y values from x values<br>
 * (http://introcs.cs.princeton.edu/java/97data/LinearRegression.java.html)
 * <b>Syntax</b>:<br>
 * <b>SLOPE</b>(<b>arrayX</b>, <b>arrayY</b>)
 */
public final class Slope extends Fixed2ArgFunction {

    private final LinearRegressionFunction func;
    public Slope() {
        func = new LinearRegressionFunction(FUNCTION.SLOPE);
    }

    @Override
    public ValueEval evaluate(int srcRowIndex, int srcColumnIndex,
            ValueEval arg0, ValueEval arg1) {
        return func.evaluate(srcRowIndex, srcColumnIndex, arg0, arg1);
    }
}

