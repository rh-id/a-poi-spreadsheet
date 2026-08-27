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
// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.PackagePart;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.ExceptionUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Methods that wrap {@link XSSFWorkbook} parsing functionality.
 * One difference is that the methods in this class try to
 * throw {@link XSSFReadException} or {@link IOException} instead of {@link RuntimeException}.
 * You can still get an {@link Error}s like an {@link OutOfMemoryError}.
 *
 * @since POI 5.5.0
 */
public final class XSSFParser {

    private XSSFParser() {
        // Prevent instantiation
    }

    /**
     * Parse the given InputStream and return a new {@link XSSFWorkbook} instance.
     *
     * @param stream the data to parse (will be closed after parsing)
     * @return a new {@link XSSFWorkbook} instance
     * @throws XSSFReadException if an error occurs while reading the file
     * @throws IOException if an I/O error occurs while reading the file
     */
    public static XSSFWorkbook parse(InputStream stream) throws XSSFReadException, IOException {
        try {
            return new XSSFWorkbook(stream);
        } catch (IOException e) {
            throw e;
        } catch (Error | RuntimeException e) {
            if (ExceptionUtil.isFatal(e)) {
                throw e;
            }
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        } catch (Exception e) {
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        }
    }

    /**
     * Parse the given InputStream and return a new {@link XSSFWorkbook} instance.
     *
     * @param stream the data to parse
     * @param closeStream whether to close the InputStream after parsing
     * @return a new {@link XSSFWorkbook} instance
     * @throws XSSFReadException if an error occurs while reading the file
     * @throws IOException if an I/O error occurs while reading the file
     */
    public static XSSFWorkbook parse(InputStream stream, boolean closeStream) throws XSSFReadException, IOException {
        try {
            return new XSSFWorkbook(stream, closeStream);
        } catch (IOException e) {
            throw e;
        } catch (Error | RuntimeException e) {
            if (ExceptionUtil.isFatal(e)) {
                throw e;
            }
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        } catch (Exception e) {
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        }
    }

    /**
     * Parse the given InputStream and return a new {@link XSSFWorkbook} instance.
     *
     * @param file to parse
     * @return a new {@link XSSFWorkbook} instance
     * @throws XSSFReadException if an error occurs while reading the file
     * @throws IOException if an I/O error occurs while reading the file
     */
    public static XSSFWorkbook parse(File file) throws XSSFReadException, IOException {
        try {
            return new XSSFWorkbook(file);
        } catch (IOException e) {
            throw e;
        } catch (Error | RuntimeException e) {
            if (ExceptionUtil.isFatal(e)) {
                throw e;
            }
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        } catch (Exception e) {
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        }
    }

    /**
     * Parse the given {@link OPCPackage} and return a new {@link XSSFWorkbook} instance.
     *
     * @param pkg to parse
     * @return a new {@link XSSFWorkbook} instance
     * @throws XSSFReadException if an error occurs while reading the file
     * @throws IOException if an I/O error occurs while reading the file
     */
    public static XSSFWorkbook parse(OPCPackage pkg) throws XSSFReadException, IOException {
        try {
            return new XSSFWorkbook(pkg);
        } catch (IOException e) {
            throw e;
        } catch (Error | RuntimeException e) {
            if (ExceptionUtil.isFatal(e)) {
                throw e;
            }
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        } catch (Exception e) {
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        }
    }

    /**
     * Parse the given {@link PackagePart} and return a new {@link XSSFWorkbook} instance.
     *
     * @param packagePart to parse
     * @return a new {@link XSSFWorkbook} instance
     * @throws XSSFReadException if an error occurs while reading the file
     * @throws IOException if an I/O error occurs while reading the file
     */
    public static XSSFWorkbook parse(PackagePart packagePart) throws XSSFReadException, IOException {
        try {
            return new XSSFWorkbook(packagePart);
        } catch (IOException e) {
            throw e;
        } catch (Error | RuntimeException e) {
            if (ExceptionUtil.isFatal(e)) {
                throw e;
            }
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        } catch (Exception e) {
            throw new XSSFReadException("Exception reading XSSFWorkbook", e);
        }
    }
}
