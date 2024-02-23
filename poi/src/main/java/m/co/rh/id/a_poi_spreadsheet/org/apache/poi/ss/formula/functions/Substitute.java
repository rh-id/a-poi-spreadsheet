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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.functions;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.ErrorEval;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.EvaluationException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.StringEval;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.eval.ValueEval;

/**
 * An implementation of the SUBSTITUTE function:<p>
 * Substitutes text in a text string with new text, some number of times.
 */
public final class Substitute extends Var3or4ArgFunction {

    public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
            ValueEval arg2) {
        String result;
        try {
            String oldStr = TextFunction.evaluateStringArg(arg0, srcRowIndex, srcColumnIndex);
            String searchStr = TextFunction.evaluateStringArg(arg1, srcRowIndex, srcColumnIndex);
            String newStr = TextFunction.evaluateStringArg(arg2, srcRowIndex, srcColumnIndex);

            result = replaceAllOccurrences(oldStr, searchStr, newStr);
        } catch (EvaluationException e) {
            return e.getErrorEval();
        }
        return new StringEval(result);
    }

    public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
            ValueEval arg2, ValueEval arg3) {
        String result;
        try {
            String oldStr = TextFunction.evaluateStringArg(arg0, srcRowIndex, srcColumnIndex);
            String searchStr = TextFunction.evaluateStringArg(arg1, srcRowIndex, srcColumnIndex);
            String newStr = TextFunction.evaluateStringArg(arg2, srcRowIndex, srcColumnIndex);

            int instanceNumber = TextFunction.evaluateIntArg(arg3, srcRowIndex, srcColumnIndex);
            if (instanceNumber < 1) {
                return ErrorEval.VALUE_INVALID;
            }
            result = replaceOneOccurrence(oldStr, searchStr, newStr, instanceNumber);
        } catch (EvaluationException e) {
            return e.getErrorEval();
        }
        return new StringEval(result);
    }

    private static String replaceAllOccurrences(String oldStr, String searchStr, String newStr) {
        // avoid endless loop when searching for nothing
        if (searchStr.length() < 1) {
            return oldStr;
        }

        StringBuilder sb = new StringBuilder();
        int startIndex = 0;
        while (true) {
            int nextMatch = oldStr.indexOf(searchStr, startIndex);
            if (nextMatch < 0) {
                // store everything from end of last match to end of string
                sb.append(oldStr.substring(startIndex));
                return sb.toString();
            }
            // store everything from end of last match to start of this match
            sb.append(oldStr, startIndex, nextMatch);
            sb.append(newStr);
            startIndex = nextMatch + searchStr.length();
        }
    }

    private static String replaceOneOccurrence(String oldStr, String searchStr, String newStr, int instanceNumber) {
        // avoid endless loop when searching for nothing
        if (searchStr.length() < 1) {
            return oldStr;
        }
        int startIndex = 0;
        int count=0;
        while (true) {
            int nextMatch = oldStr.indexOf(searchStr, startIndex);
            if (nextMatch < 0) {
                // not enough occurrences found - leave unchanged
                return oldStr;
            }
            count++;
            if (count == instanceNumber) {
                return oldStr.substring(0, nextMatch) +
                        newStr +
                        oldStr.substring(nextMatch + searchStr.length());
            }
            startIndex = nextMatch + searchStr.length();
        }
    }
}
