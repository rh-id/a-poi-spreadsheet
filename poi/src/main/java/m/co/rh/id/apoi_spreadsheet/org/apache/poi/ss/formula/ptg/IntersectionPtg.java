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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.ptg;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LittleEndianOutput;

public final class IntersectionPtg extends OperationPtg {
    public static final byte sid = 0x0f;

    public static final IntersectionPtg instance = new IntersectionPtg();

    private IntersectionPtg() {
        // enforce singleton
    }

    @Override
    public final boolean isBaseToken() {
        return true;
    }

    @Override
    public byte getSid() {
        return sid;
    }

    public int getSize() {
        return 1;
    }

    public void write(LittleEndianOutput out) {
        out.writeByte(sid + getPtgClass());
    }

    public String toFormulaString() {
        return " ";
    }

    public String toFormulaString(String[] operands) {
        return operands[0] + " " + operands[1];
    }

    public int getNumberOfOperands() {
        return 2;
    }

    @Override
    public IntersectionPtg copy() {
        return instance;
    }
}
