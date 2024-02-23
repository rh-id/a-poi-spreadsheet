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
package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.hpsf;

import android.util.Log;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.CodePageUtil;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.IOUtils;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndian;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianByteArrayInputStream;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianConsts;


@Internal
public class CodePageString {
    //arbitrarily selected; may need to increase
    private static final int DEFAULT_MAX_RECORD_LENGTH = 100_000;
    private static int MAX_RECORD_LENGTH = DEFAULT_MAX_RECORD_LENGTH;

    private static final String TAG = "CodePageString";

    private byte[] _value;

    /**
     * @param length the max record length allowed for CodePageString
     */
    public static void setMaxRecordLength(int length) {
        MAX_RECORD_LENGTH = length;
    }

    /**
     * @return the max record length allowed for CodePageString
     */
    public static int getMaxRecordLength() {
        return MAX_RECORD_LENGTH;
    }

    public void read(LittleEndianByteArrayInputStream lei) {
        int offset = lei.getReadIndex();
        int size = lei.readInt();
        _value = IOUtils.safelyAllocate(size, MAX_RECORD_LENGTH);
        if (size == 0) {
            return;
        }

        // If Size is zero, this field MUST be zero bytes in length. If Size is
        // nonzero and the CodePage property set's CodePage property has the value CP_WINUNICODE
        // (0x04B0), then the value MUST be a null-terminated array of 16-bit Unicode characters,
        // followed by zero padding to a multiple of 4 bytes. If Size is nonzero and the property set's
        // CodePage property has any other value, it MUST be a null-terminated array of 8-bit characters
        // from the code page identified by the CodePage property, followed by zero padding to a
        // multiple of 4 bytes. The string represented by this field MAY contain embedded or additional
        // trailing null characters and an OLEPS implementation MUST be able to handle such strings.

        lei.readFully(_value);
        if (_value[size - 1] != 0) {
            // TODO Some files, such as TestVisioWithCodepage.vsd, are currently
            // triggering this for values that don't look like codepages
            // See Bug #52258 for details
            Log.w(TAG, String.format("CodePageString started at offset #%d is not NULL-terminated", offset));
        }

        TypedPropertyValue.skipPadding(lei);
    }

    public String getJavaValue(int codepage) throws UnsupportedEncodingException {
        int cp = (codepage == -1) ? Property.DEFAULT_CODEPAGE : codepage;
        String result = CodePageUtil.getStringFromCodePage(_value, cp);


        final int terminator = result.indexOf('\0');
        if (terminator == -1) {
            Log.w(TAG, "String terminator (\\0) for CodePageString property value not found. " +
                    "Continue without trimming and hope for the best.");
            return result;
        }
        if (terminator != result.length() - 1) {
            Log.d(TAG, "String terminator (\\0) for CodePageString property value occurred before the end of " +
                    "string. Trimming and hope for the best.");
        }
        return result.substring(0, terminator);
    }

    public int getSize() {
        return LittleEndianConsts.INT_SIZE + _value.length;
    }

    public void setJavaValue(String string, int codepage) throws UnsupportedEncodingException {
        int cp = (codepage == -1) ? Property.DEFAULT_CODEPAGE : codepage;
        _value = CodePageUtil.getBytesInCodePage(string + "\0", cp);
    }

    public int write(OutputStream out) throws IOException {
        LittleEndian.putUInt(_value.length, out);
        out.write(_value);
        return LittleEndianConsts.INT_SIZE + _value.length;
    }
}
