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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.chart;

import java.util.Map;
import java.util.function.Supplier;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.HSSFRecordTypes;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.RecordInputStream;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.StandardRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.GenericRecordUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LittleEndianOutput;

/**
 * SERIESLIST (0x1016)<p>
 *
 * The series list record defines the series displayed as an overlay to the main chart record.<p>
 *
 * (As with all chart related records, documentation is lacking.
 * See {@link ChartRecord} for more details)
 */
public final class SeriesListRecord extends StandardRecord {
    public static final short sid = 0x1016;
    private short[] field_1_seriesNumbers;

    public SeriesListRecord(SeriesListRecord other) {
        super(other);
        field_1_seriesNumbers = (other.field_1_seriesNumbers == null) ? null : other.field_1_seriesNumbers.clone();
    }

    public SeriesListRecord(short[] seriesNumbers) {
        field_1_seriesNumbers = (seriesNumbers == null) ? null : seriesNumbers.clone();
    }

    public SeriesListRecord(RecordInputStream in) {
        int nItems = in.readUShort();
        short[] ss = new short[nItems];
        for (int i = 0; i < nItems; i++) {
            ss[i] = in.readShort();

        }
        field_1_seriesNumbers = ss;
    }

    public void serialize(LittleEndianOutput out) {

        int nItems = field_1_seriesNumbers.length;
        out.writeShort(nItems);
        for (short field_1_seriesNumber : field_1_seriesNumbers) {
            out.writeShort(field_1_seriesNumber);
        }
    }

    protected int getDataSize() {
        return field_1_seriesNumbers.length * 2 + 2;
    }

    public short getSid() {
        return sid;
    }

    @Override
    public SeriesListRecord copy() {
        return new SeriesListRecord(this);
    }

    /**
     * Get the series numbers field for the SeriesList record.
     */
    public short[] getSeriesNumbers() {
        return field_1_seriesNumbers;
    }

    @Override
    public HSSFRecordTypes getGenericRecordType() {
        return HSSFRecordTypes.SERIES_LIST;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return GenericRecordUtil.getGenericProperties("seriesNumbers", this::getSeriesNumbers);
    }
}
