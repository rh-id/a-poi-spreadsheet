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

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FormulaError;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.GenericRecordUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LittleEndianOutput;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.RecordFormatException;

/**
 * Creates new BoolErrRecord. (0x0205)
 */
public final class BoolErrRecord extends CellRecord {
    public static final short sid = 0x0205;
    private int _value;
    /**
     * If <code>true</code>, this record represents an error cell value,
     * otherwise this record represents a boolean cell value
     */
    private boolean _isError;

    /** Creates new BoolErrRecord */
    public BoolErrRecord() {}

    public BoolErrRecord(BoolErrRecord other) {
        super(other);
        _value = other._value;
        _isError = other._isError;
    }

    /**
     * @param in the RecordInputstream to read the record from
     */
    public BoolErrRecord(RecordInputStream in) {
        super(in);
        switch (in.remaining()) {
            case 2:
                _value = in.readByte();
                break;
            case 3:
                _value = in.readUShort();
                break;
            default:
                throw new RecordFormatException("Unexpected size ("
                        + in.remaining() + ") for BOOLERR record.");
        }
        int flag = in.readUByte();
        switch (flag) {
            case 0:
                _isError = false;
                break;
            case 1:
                _isError = true;
                break;
            default:
                throw new RecordFormatException("Unexpected isError flag ("
                        + flag + ") for BOOLERR record.");
        }
    }

    /**
     * set the boolean value for the cell
     *
     * @param value   representing the boolean value
     */
    public void setValue(boolean value) {
        _value = value ? 1 : 0;
        _isError = false;
    }

    /**
     * set the error value for the cell. See {@link FormulaError} for valid codes.
     *
     * @param value     error representing the error value
     *                  this value can only be 0,7,15,23,29,36 or 42
     *                  see bugzilla bug 16560 for an explanation
     */
    public void setValue(byte value) {
        setValue(FormulaError.forInt(value));
    }

    /**
     * set the error value for the cell
     *
     * @param value     error representing the error value
     *                  this value can only be 0,7,15,23,29,36 or 42
     *                  see bugzilla bug 16560 for an explanation
     */
    public void setValue(FormulaError value) {
        switch(value) {
            case NULL:
            case DIV0:
            case VALUE:
            case REF:
            case NAME:
            case NUM:
            case NA:
                _value = value.getCode();
                _isError = true;
                return;
            default:
                throw new IllegalArgumentException("Error Value can only be 0,7,15,23,29,36 or 42. It cannot be "+value.getCode()+" ("+value+")");
        }
    }

    /**
     * get the value for the cell
     *
     * @return boolean representing the boolean value
     */
    public boolean getBooleanValue() {
        return _value != 0;
    }

    /**
     * get the error value for the cell
     *
     * @return byte representing the error value
     */
    public byte getErrorValue() {
        return (byte)_value;
    }

    /**
     * Indicates whether the call holds a boolean value
     *
     * @return boolean true if the cell holds a boolean value
     */
    public boolean isBoolean() {
        return !_isError;
    }

    /**
     * Indicates whether the call holds an error value
     *
     * @return boolean true if the cell holds an error value
     */
    public boolean isError() {
        return _isError;
    }

    @Override
    protected String getRecordName() {
        return "BOOLERR";
    }

    @Override
    protected void serializeValue(LittleEndianOutput out) {
        out.writeByte(_value);
        out.writeByte(_isError ? 1 : 0);
    }

    @Override
    protected int getValueDataSize() {
        return 2;
    }

    public short getSid() {
        return sid;
    }

    @Override
    public BoolErrRecord copy() {
        return new BoolErrRecord(this);
    }

    @Override
    public HSSFRecordTypes getGenericRecordType() {
        return HSSFRecordTypes.BOOL_ERR;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return GenericRecordUtil.getGenericProperties(
            "base", super::getGenericProperties,
            "isBoolean", this::isBoolean,
            "booleanVal", this::getBooleanValue,
            "isError", this::isError,
            "errorVal", this::getErrorValue,
            "errorTxt", () -> isError() ? FormulaError.forInt(getErrorValue()).getString() : null
        );
    }
}
