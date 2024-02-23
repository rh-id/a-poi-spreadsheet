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

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.FunctionNameEval;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.NotImplementedFunctionException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.ValueEval;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.functions.ArrayFunction;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.functions.FreeRefFunction;

/**
 *
 * Common entry point for all user-defined (non-built-in) functions (where
 * {@code AbstractFunctionPtg.field_2_fnc_index} == 255)
 */
final class UserDefinedFunction implements FreeRefFunction {

    public static final FreeRefFunction instance = new UserDefinedFunction();

    private UserDefinedFunction() {
        // enforce singleton
    }

    /**
     * @param args the pre-evaluated arguments for this function. args is never {@code null},
     *             nor are any of its elements.
     * @param ec primarily used to identify the source cell containing the formula being evaluated.
     *             may also be used to dynamically create reference evals.
     * @return value
     * @throws IllegalStateException if first arg is not a {@link FunctionNameEval}
     * @throws NotImplementedFunctionException if function is not implemented
     */
    @Override
    public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
        int nIncomingArgs = args.length;
        if (nIncomingArgs < 1) {
            throw new IllegalStateException("function name argument missing");
        }

        ValueEval nameArg = args[0];
        String functionName;
        if (nameArg instanceof FunctionNameEval) {
            functionName = ((FunctionNameEval) nameArg).getFunctionName();
        } else {
            throw new IllegalStateException("First argument should be a NameEval, but got ("
                    + nameArg.getClass().getName() + ")");
        }
        FreeRefFunction targetFunc = ec.findUserDefinedFunction(functionName);
        if (targetFunc == null) {
            throw new NotImplementedFunctionException(functionName);
        }
        int nOutGoingArgs = nIncomingArgs - 1;
        ValueEval[] outGoingArgs = new ValueEval[nOutGoingArgs];
        System.arraycopy(args, 1, outGoingArgs, 0, nOutGoingArgs);
        if (targetFunc instanceof ArrayFunction) {
            ArrayFunction func = (ArrayFunction) targetFunc;
            ValueEval eval = OperationEvaluatorFactory.evaluateArrayFunction(func, outGoingArgs, ec);
            if (eval != null) {
                return eval;
            }
        }
        return targetFunc.evaluate(outGoingArgs, ec);
    }
}