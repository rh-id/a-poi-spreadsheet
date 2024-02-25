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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.internal;

import android.util.Log;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;

import m.co.rh.id.apoi_spreadsheet.base.util.TempFile;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.PackagePart;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.PackagePartName;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.internal.marshallers.ZipPartMarshaller;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.IOUtils;

/**
 * (Experimental) Temp File version of a package part.
 *
 * @since POI 5.1.0
 */
@Beta
public final class TempFilePackagePart extends PackagePart {
    private static final String TAG = "TempFilePackagePart";

    /**
     * Storage for the part data.
     */
    private File tempFile;

    /**
     * Constructor.
     *
     * @param pack        The owner package.
     * @param partName    The part name.
     * @param contentType The content type.
     * @throws InvalidFormatException If the specified URI is not OPC compliant.
     * @throws IOException            If temp file cannot be created.
     */
    public TempFilePackagePart(OPCPackage pack, PackagePartName partName,
                               String contentType) throws InvalidFormatException, IOException {
        this(pack, partName, contentType, true);
    }

    /**
     * Constructor.
     *
     * @param pack              The owner package.
     * @param partName          The part name.
     * @param contentType       The content type.
     * @param loadRelationships Specify if the relationships will be loaded.
     * @throws InvalidFormatException If the specified URI is not OPC compliant.
     * @throws IOException            If temp file cannot be created.
     */
    public TempFilePackagePart(OPCPackage pack, PackagePartName partName,
                               String contentType, boolean loadRelationships)
            throws InvalidFormatException, IOException {
        super(pack, partName, new ContentType(contentType), loadRelationships);
        tempFile = TempFile.createTempFile("poi-package-part", ".tmp");
    }

    @Override
    protected InputStream getInputStreamImpl() throws IOException {
        return Files.newInputStream(tempFile.toPath());
    }

    @Override
    protected OutputStream getOutputStreamImpl() throws IOException {
        return Files.newOutputStream(tempFile.toPath());
    }

    @Override
    public long getSize() {
        return tempFile.length();
    }

    @Override
    public void clear() {
        try (OutputStream os = getOutputStreamImpl()) {
            os.write(new byte[0]);
        } catch (IOException e) {
            Log.w(TAG, "Failed to clear data in temp file", e);
        }
    }

    @Override
    public boolean save(OutputStream os) throws OpenXML4JException {
        return new ZipPartMarshaller().marshall(this, os);
    }

    @Override
    public boolean load(InputStream is) throws InvalidFormatException {
        try (OutputStream os = getOutputStreamImpl()) {
            IOUtils.copy(is, os);
        } catch (IOException e) {
            throw new InvalidFormatException(e.getMessage(), e);
        }

        // All done
        return true;
    }

    @Override
    public void close() {
        if (!tempFile.delete()) {
            Log.i(TAG, "Failed to delete temp file; may already have been closed and deleted");
        }
    }

    @Override
    public void flush() {
        // Do nothing
    }
}
