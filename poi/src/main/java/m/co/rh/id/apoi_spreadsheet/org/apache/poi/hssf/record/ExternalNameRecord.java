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

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.Formula;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.constant.ConstantValueParser;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.formula.ptg.Ptg;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.GenericRecordUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.LittleEndianOutput;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.StringUtil;

/**
 * EXTERNALNAME (0x0023)
 */
public final class ExternalNameRecord extends StandardRecord {

    public static final short sid = 0x0023; // as per BIFF8. (some old versions used 0x223)

    private static final int OPT_BUILTIN_NAME          = 0x0001;
    private static final int OPT_AUTOMATIC_LINK        = 0x0002; // m$ doc calls this fWantAdvise
    private static final int OPT_PICTURE_LINK          = 0x0004;
    private static final int OPT_STD_DOCUMENT_NAME     = 0x0008; //fOle
    private static final int OPT_OLE_LINK              = 0x0010; //fOleLink
//  private static final int OPT_CLIP_FORMAT_MASK      = 0x7FE0;
    private static final int OPT_ICONIFIED_PICTURE_LINK= 0x8000;

    private static final int[] OPTION_FLAGS = {
        OPT_BUILTIN_NAME,OPT_AUTOMATIC_LINK,OPT_PICTURE_LINK,OPT_STD_DOCUMENT_NAME,OPT_OLE_LINK,OPT_ICONIFIED_PICTURE_LINK};
    private static final String[] OPTION_NAMES = {
        "BUILTIN_NAME","AUTOMATIC_LINK","PICTURE_LINK","STD_DOCUMENT_NAME","OLE_LINK","ICONIFIED_PICTURE_LINK"};



    private short field_1_option_flag;
    private short field_2_ixals;
    private short field_3_not_used;
    private String field_4_name;
    private Formula field_5_name_definition;

    /**
     * 'rgoper' / 'Last received results of the DDE link'
     * (seems to be only applicable to DDE links)<br>
     * Logically this is a 2-D array, which has been flattened into 1-D array here.
     */
    private Object[] _ddeValues;
    /**
     * (logical) number of columns in the {@link #_ddeValues} array
     */
    private int _nColumns;
    /**
     * (logical) number of rows in the {@link #_ddeValues} array
     */
    private int _nRows;

    public ExternalNameRecord() {
        field_2_ixals = 0;
    }

    public ExternalNameRecord(ExternalNameRecord other) {
        super(other);
        field_1_option_flag = other.field_1_option_flag;
        field_2_ixals = other.field_2_ixals;
        field_3_not_used = other.field_3_not_used;
        field_4_name = other.field_4_name;
        field_5_name_definition = (other.field_5_name_definition == null) ? null : other.field_5_name_definition.copy();
        _ddeValues = (other._ddeValues == null) ? null : other._ddeValues.clone();
        _nColumns = other._nColumns;
        _nRows = other._nRows;
    }

    public ExternalNameRecord(RecordInputStream in) {
        field_1_option_flag = in.readShort();
        field_2_ixals       = in.readShort();
        field_3_not_used    = in.readShort();

        int numChars = in.readUByte();
        field_4_name = StringUtil.readUnicodeString(in, numChars);

        // the record body can take different forms.
        // The form is dictated by the values of 3-th and 4-th bits in field_1_option_flag
        if(!isOLELink() && !isStdDocumentNameIdentifier()){
            // another switch: the fWantAdvise bit specifies whether the body describes
            // an external defined name or a DDE data item
            if(isAutomaticLink()){
                if(in.available() > 0) {
                    //body specifies DDE data item
                    int nColumns = in.readUByte() + 1;
                    int nRows = in.readShort() + 1;

                    int totalCount = nRows * nColumns;
                    _ddeValues = ConstantValueParser.parse(in, totalCount);
                    _nColumns = nColumns;
                    _nRows = nRows;
                }
            } else {
                //body specifies an external defined name
                int formulaLen = in.readUShort();
                field_5_name_definition = Formula.read(formulaLen, in);
            }
        }
    }

    /**
     * @return {@code true} if the name is a built-in name
     */
    public boolean isBuiltInName() {
        return (field_1_option_flag & OPT_BUILTIN_NAME) != 0;
    }
    /**
     * For OLE and DDE, links can be either 'automatic' or 'manual'
     *
     * @return {@code true} if this is an automatic link
     */
    public boolean isAutomaticLink() {
        return (field_1_option_flag & OPT_AUTOMATIC_LINK) != 0;
    }
    /**
     * only for OLE and DDE
     *
     * @return {@code true} if this is a picture link
     */
    public boolean isPicureLink() {
        return (field_1_option_flag & OPT_PICTURE_LINK) != 0;
    }
    /**
     * DDE links only. If <code>true</code>, this denotes the 'StdDocumentName'
     *
     * @return {@code true} if this denotes the 'StdDocumentName'
     */
    public boolean isStdDocumentNameIdentifier() {
        return (field_1_option_flag & OPT_STD_DOCUMENT_NAME) != 0;
    }
    public boolean isOLELink() {
        return (field_1_option_flag & OPT_OLE_LINK) != 0;
    }
    public boolean isIconifiedPictureLink() {
        return (field_1_option_flag & OPT_ICONIFIED_PICTURE_LINK) != 0;
    }
    /**
     * @return the standard String representation of this name
     */
    public String getText() {
        return field_4_name;
    }

    public void setText(String str) {
        field_4_name = str;
    }

    /**
     * If this is a local name, then this is the (1 based)
     *  index of the name of the Sheet this refers to, as
     *  defined in the preceding {@link SupBookRecord}.
     * If it isn't a local name, then it must be zero.
     *
     * @return the index of the name of the Sheet this refers to
     */
    public short getIx() {
       return field_2_ixals;
    }

    public void setIx(short ix) {
        field_2_ixals = ix;
    }

    public Ptg[] getParsedExpression() {
        return Formula.getTokens(field_5_name_definition);
    }
    public void setParsedExpression(Ptg[] ptgs) {
        field_5_name_definition = Formula.create(ptgs);
    }


    @Override
    protected int getDataSize(){
        int result = 2 + 4;  // short and int
        result += StringUtil.getEncodedSize(field_4_name) - 1; //size is byte, not short

        if(!isOLELink() && !isStdDocumentNameIdentifier()){
            if(isAutomaticLink()){
                if(_ddeValues != null) {
                    result += 3; // byte, short
                    result += ConstantValueParser.getEncodedSize(_ddeValues);
                }
            } else {
                result += field_5_name_definition.getEncodedSize();
            }
        }
        return result;
    }

    @Override
    public void serialize(LittleEndianOutput out) {
        out.writeShort(field_1_option_flag);
        out.writeShort(field_2_ixals);
        out.writeShort(field_3_not_used);

        out.writeByte(field_4_name.length());
        StringUtil.writeUnicodeStringFlagAndData(out, field_4_name);

        if(!isOLELink() && !isStdDocumentNameIdentifier()){
            if(isAutomaticLink()){
                if(_ddeValues != null) {
                    out.writeByte(_nColumns-1);
                    out.writeShort(_nRows-1);
                    ConstantValueParser.encode(out, _ddeValues);
                }
            } else {
                field_5_name_definition.serialize(out);
            }
        }
    }

    @Override
    public short getSid() {
        return sid;
    }

    @Override
    public ExternalNameRecord copy() {
        return new ExternalNameRecord(this);
    }

    @Override
    public HSSFRecordTypes getGenericRecordType() {
        return HSSFRecordTypes.EXTERNAL_NAME;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return GenericRecordUtil.getGenericProperties(
            "options", GenericRecordUtil.getBitsAsString(() -> field_1_option_flag, OPTION_FLAGS, OPTION_NAMES),
            "ix", this::getIx,
            "name", this::getText,
            "nameDefinition", (field_5_name_definition == null ? () -> null : field_5_name_definition::getTokens)
        );

    }
}
