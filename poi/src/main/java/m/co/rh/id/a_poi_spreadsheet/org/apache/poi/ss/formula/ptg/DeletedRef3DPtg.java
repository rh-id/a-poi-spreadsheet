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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.ptg;


import java.util.Map;
import java.util.function.Supplier;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.formula.WorkbookDependentFormula;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.FormulaError;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.GenericRecordUtil;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianInput;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianOutput;

/**
 * Deleted Reference 3D Ptg<p>
 * Defined a cell in extern sheet.
 *
 * @since 1.0-pre
 */
public final class DeletedRef3DPtg extends OperandPtg implements WorkbookDependentFormula {
    public static final byte sid  = 0x3c;
    private final int field_1_index_extern_sheet;
    private final int unused1;

    /** Creates new DeletedRef3DPtg */
    public DeletedRef3DPtg(LittleEndianInput in)  {
        field_1_index_extern_sheet = in.readUShort();
        unused1 = in.readInt();
    }

    public DeletedRef3DPtg(int externSheetIndex) {
        field_1_index_extern_sheet = externSheetIndex;
        unused1 = 0;
    }

    public String toFormulaString(FormulaRenderingWorkbook book) {
        return ExternSheetNameResolver.prependSheetName(book, field_1_index_extern_sheet, FormulaError.REF.getString());
    }
    public String toFormulaString() {
        throw new IllegalStateException("3D references need a workbook to determine formula text");
    }
    @Override
    public byte getDefaultOperandClass() {
        return Ptg.CLASS_REF;
    }

    @Override
    public byte getSid() {
        return sid;
    }

    public int getSize() {
        return 7;
    }
    public void write(LittleEndianOutput out) {
        out.writeByte(sid + getPtgClass());
        out.writeShort(field_1_index_extern_sheet);
        out.writeInt(unused1);
    }

    @Override
    public DeletedRef3DPtg copy() {
        // immutable
        return this;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return GenericRecordUtil.getGenericProperties(
            "externSheetIndex", () -> field_1_index_extern_sheet,
            "unused1", () -> unused1
        );
    }
}