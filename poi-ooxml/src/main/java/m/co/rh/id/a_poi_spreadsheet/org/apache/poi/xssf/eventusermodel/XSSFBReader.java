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
package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.eventusermodel;

import android.util.Log;

import com.zaxxer.sparsebits.SparseBitSet;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.PackagePart;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.PackagePartName;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.PackageRelationship;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.PackagingURIHelper;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndian;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.LittleEndianConsts;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.StringUtil;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBCommentsTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBParseException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBParser;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBRecordType;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBRelation;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBStylesTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBUtils;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.model.CommentsTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Reader for xlsb files.
 *
 * @since 3.16-beta3
 */
public class XSSFBReader extends XSSFReader {

    private static final String TAG = "XSSFBReader";
    private static final Set<String> WORKSHEET_RELS =
            Collections.unmodifiableSet(new HashSet<>(
                    Arrays.asList(
                            XSSFRelation.WORKSHEET.getRelation(),
                            XSSFRelation.CHARTSHEET.getRelation(),
                            XSSFRelation.MACRO_SHEET_XML.getRelation(),
                            XSSFRelation.INTL_MACRO_SHEET_XML.getRelation(),
                            XSSFRelation.DIALOG_SHEET_BIN.getRelation()
                    )
            ));

    /**
     * Creates a new XSSFReader, for the given package
     *
     * @param pkg opc package
     * @throws OpenXML4JException if the package data format is invalid
     * @throws IOException        if there is an I/O issue reading the data
     */
    public XSSFBReader(OPCPackage pkg) throws IOException, OpenXML4JException {
        super(pkg);
    }

    /**
     * In Excel 2013, the absolute path where the file was last saved may be stored in
     * the {@link XSSFBRecordType#BrtAbsPath15} record.  The equivalent in ooxml is
     * &lt;x15ac:absPath&gt;.
     *
     * @return absolute path or <code>null</code> if it could not be found.
     * @throws IOException when there's a problem with the workbook part's stream
     */
    public String getAbsPathMetadata() throws IOException {
        try (InputStream is = workbookPart.getInputStream()) {
            PathExtractor p = new PathExtractor(is);
            p.parse();
            return p.getPath();
        }
    }

    /**
     * Returns an Iterator which will let you get at all the
     * different Sheets in turn.
     * Each sheet's InputStream is only opened when fetched
     * from the Iterator. It's up to you to close the
     * InputStreams when done with each one.
     *
     * @return iterator of {@link InputStream}s
     * @throws InvalidFormatException if the sheet data format is invalid
     * @throws IOException            if there is an I/O issue reading the data
     */
    @Override
    public Iterator<InputStream> getSheetsData() throws IOException, InvalidFormatException {
        return new SheetIterator(workbookPart);
    }

    public XSSFBStylesTable getXSSFBStylesTable() throws IOException {
        ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFBRelation.STYLES_BINARY.getContentType());
        if (parts.isEmpty()) return null;

        // Create the Styles Table, and associate the Themes if present
        try (InputStream stream = parts.get(0).getInputStream()) {
            return new XSSFBStylesTable(stream);
        }
    }

    public static class SheetIterator extends XSSFReader.SheetIterator {

        /**
         * Construct a new SheetIterator
         *
         * @param wb package part holding workbook.xml
         * @throws InvalidFormatException if the sheet data format is invalid
         * @throws IOException            if there is an I/O issue reading the data
         */
        private SheetIterator(PackagePart wb) throws IOException, InvalidFormatException {
            super(wb);
        }

        @Override
        protected Set<String> getSheetRelationships() {
            return WORKSHEET_RELS;
        }

        @Override
        protected Iterator<XSSFSheetRef> createSheetIteratorFromWB(PackagePart wb) throws IOException {
            try (InputStream stream = wb.getInputStream()) {
                SheetRefLoader sheetRefLoader = new SheetRefLoader(stream);
                sheetRefLoader.parse();
                return sheetRefLoader.getSheets().iterator();
            }
        }

        /**
         * Not supported by XSSFBReader's SheetIterator.
         * Please use {@link #getXSSFBSheetComments()} instead.
         *
         * @return nothing, always throws IllegalArgumentException!
         */
        @Override
        public CommentsTable getSheetComments() {
            throw new IllegalArgumentException("Please use getXSSFBSheetComments");
        }

        public XSSFBCommentsTable getXSSFBSheetComments() {
            PackagePart sheetPkg = getSheetPart();

            // Do we have a comments relationship? (Only ever one if so)
            try {
                PackageRelationshipCollection commentsList =
                        sheetPkg.getRelationshipsByType(XSSFRelation.SHEET_COMMENTS.getRelation());
                if (!commentsList.isEmpty()) {
                    PackageRelationship comments = commentsList.getRelationship(0);
                    if (comments == null || comments.getTargetURI() == null) {
                        return null;
                    }
                    PackagePartName commentsName = PackagingURIHelper.createPartName(comments.getTargetURI());
                    PackagePart commentsPart = sheetPkg.getPackage().getPart(commentsName);
                    try (InputStream stream = commentsPart.getInputStream()) {
                        return new XSSFBCommentsTable(stream);
                    }
                }
            } catch (InvalidFormatException | IOException e) {
                return null;
            }
            return null;
        }

    }


    private static class PathExtractor extends XSSFBParser {
        private static final SparseBitSet RECORDS = new SparseBitSet();

        static {
            RECORDS.set(XSSFBRecordType.BrtAbsPath15.getId());
        }

        private String path;

        public PathExtractor(InputStream is) {
            super(is, RECORDS);
        }

        @Override
        public void handleRecord(int recordType, byte[] data) throws XSSFBParseException {
            if (recordType != XSSFBRecordType.BrtAbsPath15.getId()) {
                return;
            }
            StringBuilder sb = new StringBuilder();
            XSSFBUtils.readXLWideString(data, 0, sb);
            path = sb.toString();
        }

        /**
         * @return the path if found, otherwise <code>null</code>
         */
        String getPath() {
            return path;
        }
    }

    private static class SheetRefLoader extends XSSFBParser {
        List<XSSFSheetRef> sheets = new LinkedList<>();

        private SheetRefLoader(InputStream is) {
            super(is);
        }

        @Override
        public void handleRecord(int recordType, byte[] data) throws XSSFBParseException {
            if (recordType == XSSFBRecordType.BrtBundleSh.getId()) {
                addWorksheet(data);
            }
        }

        private void addWorksheet(byte[] data) {
            //try to parse the BrtBundleSh
            //if there's an exception, catch it and
            //try to figure out if this is one of the old beta-created xlsb files
            //or if this is a general exception
            try {
                tryToAddWorksheet(data);
            } catch (XSSFBParseException e) {
                if (tryOldFormat(data)) {
                    Log.w(TAG, "This file was written with a beta version of Excel. " +
                            "POI will try to parse the file as a regular xlsb.");
                } else {
                    throw e;
                }
            }
        }

        private void tryToAddWorksheet(byte[] data) throws XSSFBParseException {
            int offset = 0;
            //this is the sheet state #2.5.142
            /*long hsShtat =*/ //noinspection ResultOfMethodCallIgnored
            LittleEndian.getUInt(data, offset);
            offset += LittleEndianConsts.INT_SIZE;

            long iTabID = LittleEndian.getUInt(data, offset);
            offset += LittleEndianConsts.INT_SIZE;
            //according to #2.4.304
            if (iTabID < 1 || iTabID > 0x0000FFFFL) {
                throw new XSSFBParseException("table id out of range: " + iTabID);
            }
            StringBuilder sb = new StringBuilder();
            offset += XSSFBUtils.readXLWideString(data, offset, sb);
            String relId = sb.toString();
            sb.setLength(0);
            /*offset +=*/
            XSSFBUtils.readXLWideString(data, offset, sb);
            String name = sb.toString();
            if (StringUtil.isNotBlank(relId)) {
                sheets.add(new XSSFSheetRef(relId, name));
            }
        }

        private boolean tryOldFormat(byte[] data) throws XSSFBParseException {
            //undocumented what is contained in these 8 bytes.
            //for the non-beta xlsb files, this would be 4, not 8.
            int offset = 8;
            long iTabID = LittleEndian.getUInt(data, offset);
            offset += LittleEndianConsts.INT_SIZE;
            if (iTabID < 1 || iTabID > 0x0000FFFFL) {
                throw new XSSFBParseException("table id out of range: " + iTabID);
            }
            StringBuilder sb = new StringBuilder();
            offset += XSSFBUtils.readXLWideString(data, offset, sb);
            String relId = sb.toString();
            sb.setLength(0);
            offset += XSSFBUtils.readXLWideString(data, offset, sb);
            String name = sb.toString();
            if (StringUtil.isNotBlank(relId)) {
                sheets.add(new XSSFSheetRef(relId, name));
            }
            return offset == data.length;
        }

        List<XSSFSheetRef> getSheets() {
            return sheets;
        }
    }
}
