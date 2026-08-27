// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.crypt.temp;

import android.util.Log;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.util.ZipEntrySource;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.IOUtils;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SXSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.streaming.SheetDataWriter;




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
        Log.i(TAG, String.format("isCompressTempFiles: %s", isCompressTempFiles()));
        Log.i(TAG, String.format("SharedStringSource: %s", getSharedStringSource()));
        return new SheetDataWriterWithDecorator();
    }
}
