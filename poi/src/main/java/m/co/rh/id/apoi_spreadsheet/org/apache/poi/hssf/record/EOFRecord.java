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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record;

import java.util.Map;
import java.util.function.Supplier;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LittleEndianOutput;

/**
 * Marks the end of records belonging to a particular object in the HSSF File
 *
 * @version 2.0-pre
 */
public final class EOFRecord extends StandardRecord {
    public static final short sid = 0x0A;
    public static final int ENCODED_SIZE = 4;

    public static final EOFRecord instance = new EOFRecord();

    private EOFRecord() {}

    /**
     * @param in unused (since this record has no data)
     */
    public EOFRecord(RecordInputStream in) {}

    public void serialize(LittleEndianOutput out) {
    }

    protected int getDataSize() {
        return ENCODED_SIZE - 4;
    }

    public short getSid()
    {
        return sid;
    }

    @Override
    public EOFRecord copy() {
      return instance;
    }

    @Override
    public HSSFRecordTypes getGenericRecordType() {
        return HSSFRecordTypes.EOF;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return null;
    }
}
