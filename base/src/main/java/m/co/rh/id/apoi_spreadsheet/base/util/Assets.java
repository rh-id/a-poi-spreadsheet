package m.co.rh.id.apoi_spreadsheet.base.util;

import android.content.Context;
import android.util.Log;

import java.io.IOException;
import java.io.InputStream;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;

/**
 * Helper to open files shipped as Android assets.
 * <p>
 * Android assets live outside the JVM classpath, so they are invisible to
 * {@link Class#getResourceAsStream(String)}. POI code ported from upstream that loads bundled
 * resources through the classpath falls back to this class instead.
 * <p>
 * Such resources are shipped by the fork's modules under the legacy org/apache/poi/... folder
 * inside each module's src/main/assets directory.
 */
public final class Assets {

    private static final String TAG = "POIAssets";

    /**
     * Opens the given asset path using the application context registered in
     * {@link POISpreadsheetContext}.
     *
     * @param path asset path relative to the assets root,
     *             e.g: org/apache/poi/ss/formula/function/functionMetadata.txt
     * @return an open {@link InputStream} for the asset, or null when no application context has
     * been set yet or the asset cannot be found
     */
    public static InputStream openOrNull(String path) {
        Context context = POISpreadsheetContext.getInstance().getAppContext();
        if (context == null) {
            return null;
        }
        try {
            return context.getAssets().open(path);
        } catch (IOException e) {
            Log.w(TAG, "Asset not found: " + path);
            return null;
        }
    }

    private Assets() {
    }
}
