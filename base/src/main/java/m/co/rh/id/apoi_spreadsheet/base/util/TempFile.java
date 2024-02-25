package m.co.rh.id.apoi_spreadsheet.base.util;

import android.content.Context;

import java.io.File;
import java.io.IOException;
import java.util.UUID;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;

public final class TempFile {

    private static final String TEMP_NAMESPACE = "m.co.rh.id.a_poi_spreadsheet";

    private static File getTempDir() {
        Context context = POISpreadsheetContext.getInstance().getAppContext();
        if (context == null) {
            throw new UnsupportedOperationException("Context has not been set");
        }
        File cacheDir = context.getCacheDir();
        String uuid = UUID.randomUUID().toString();
        File resultFile = new File(cacheDir, TEMP_NAMESPACE + "/" + uuid);
        resultFile.mkdirs();
        return resultFile;
    }

    public static File createTempFile(String prefix, String suffix) throws IOException {
        return new File(getTempDir(), prefix + suffix);
    }

    public static File createTempDirectory(String name) throws IOException {
        return new File(getTempDir(), name);
    }

    private TempFile() {
    }
}
