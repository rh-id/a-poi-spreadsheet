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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.record.pivottable;

import java.util.Map;
import java.util.function.Supplier;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.record.HSSFRecordTypes;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.record.RecordInputStream;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hssf.record.StandardRecord;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.GenericRecordUtil;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianOutput;

/**
 * SXVS - View Source (0x00E3)<br>
 */
public final class ViewSourceRecord extends StandardRecord {
    public static final short sid = 0x00E3;

    private int vs;

    public ViewSourceRecord(ViewSourceRecord other) {
        super(other);
        vs = other.vs;
    }

    public ViewSourceRecord(RecordInputStream in) {
        vs = in.readShort();
    }

    @Override
    protected void serialize(LittleEndianOutput out) {
        out.writeShort(vs);
    }

    @Override
    protected int getDataSize() {
        return 2;
    }

    @Override
    public short getSid() {
        return sid;
    }

    @Override
    public ViewSourceRecord copy() {
        return new ViewSourceRecord(this);
    }

    @Override
    public HSSFRecordTypes getGenericRecordType() {
        return HSSFRecordTypes.VIEW_SOURCE;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return GenericRecordUtil.getGenericProperties("vs", () -> vs);
    }
}