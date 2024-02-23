/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.poifs.crypt.temp;

import android.util.Log;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.util.ZipEntrySource;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.IOUtils;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFWorkbook;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.streaming.SheetDataWriter;


@Beta
public class SXSSFWorkbookWithCustomZipEntrySource extends SXSSFWorkbook {
    private static final String TAG = "SXSSFWorkbookWithCustomZipEntrySource";

    public SXSSFWorkbookWithCustomZipEntrySource() {
        super(20);
        setCompressTempFiles(true);
    }
    
    @Override
    public void write(OutputStream stream) throws IOException {
        flushSheets();
        EncryptedTempData tempData = new EncryptedTempData();
        ZipEntrySource source = null;
        try {
            try (OutputStream os = tempData.getOutputStream()) {
                getXSSFWorkbook().write(os);
            }
            // provide ZipEntrySource to poi which decrypts on the fly
            try (InputStream tempStream = tempData.getInputStream()) {
                source = AesZipFileZipEntrySource.createZipEntrySource(tempStream);
            }
            injectData(source, stream);
        } finally {
            tempData.dispose();
            IOUtils.closeQuietly(source);
        }
    }
    
    @Override
    protected SheetDataWriter createSheetDataWriter() throws IOException {
        //log values to ensure these values are accessible to subclasses
        Log.i(TAG, String.format("isCompressTempFiles: %b", isCompressTempFiles()));
        Log.i(TAG,String.format("SharedStringSource: %s", getSharedStringSource()));
        return new SheetDataWriterWithDecorator();
    }
}