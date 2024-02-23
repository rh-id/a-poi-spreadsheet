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
package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.extractor;

import android.util.Log;

import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.List;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.openxml4j.opc.OPCPackage;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.DataFormatter;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBCommentsTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBHyperlinksTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBSheetHandler;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.binary.XSSFBStylesTable;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.eventusermodel.XSSFBReader;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.model.SharedStrings;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Implementation of a text extractor or xlsb Excel
 * files that uses SAX-like binary parsing.
 *
 * @since 3.16-beta3
 */
public class XSSFBEventBasedExcelExtractor extends XSSFEventBasedExcelExtractor {

    private static final String TAG = "XSSFBEventBasedExcelExtractor";

    public static final List<XSSFRelation> SUPPORTED_TYPES = Collections.singletonList(
            XSSFRelation.XLSB_BINARY_WORKBOOK
    );

    private boolean handleHyperlinksInCells;

    public XSSFBEventBasedExcelExtractor(String path) throws XmlException, OpenXML4JException, IOException {
        super(path);
    }

    public XSSFBEventBasedExcelExtractor(OPCPackage container) throws XmlException, OpenXML4JException, IOException {
        super(container);
    }

    public void setHandleHyperlinksInCells(boolean handleHyperlinksInCells) {
        this.handleHyperlinksInCells = handleHyperlinksInCells;
    }

    /**
     * Should we return the formula itself, and not
     * the result it produces? Default is false
     * This is currently unsupported for xssfb
     */
    @Override
    public void setFormulasNotResults(boolean formulasNotResults) {
        throw new IllegalArgumentException("Not currently supported");
    }

    /**
     * Processes the given sheet
     */
    public void processSheet(
            SheetContentsHandler sheetContentsExtractor,
            XSSFBStylesTable styles,
            XSSFBCommentsTable comments,
            SharedStrings strings,
            InputStream sheetInputStream)
            throws IOException {

        DataFormatter formatter;
        if (getLocale() == null) {
            formatter = new DataFormatter();
        } else {
            formatter = new DataFormatter(getLocale());
        }

        XSSFBSheetHandler xssfbSheetHandler = new XSSFBSheetHandler(
                sheetInputStream,
                styles, comments, strings, sheetContentsExtractor, formatter, getFormulasNotResults()
        );
        xssfbSheetHandler.parse();
    }

    /**
     * Processes the file and returns the text
     */
    public String getText() {
        try {
            XSSFBSharedStringsTable strings = new XSSFBSharedStringsTable(getPackage());
            XSSFBReader xssfbReader = new XSSFBReader(getPackage());
            XSSFBStylesTable styles = xssfbReader.getXSSFBStylesTable();
            XSSFBReader.SheetIterator iter = (XSSFBReader.SheetIterator) xssfbReader.getSheetsData();

            StringBuilder text = new StringBuilder(64);
            SheetTextExtractor sheetExtractor = new SheetTextExtractor();
            XSSFBHyperlinksTable hyperlinksTable = null;
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    if (getIncludeSheetNames()) {
                        text.append(iter.getSheetName());
                        text.append('\n');
                    }
                    if (handleHyperlinksInCells) {
                        hyperlinksTable = new XSSFBHyperlinksTable(iter.getSheetPart());
                    }
                    XSSFBCommentsTable comments = getIncludeCellComments() ? iter.getXSSFBSheetComments() : null;
                    processSheet(sheetExtractor, styles, comments, strings, stream);
                    if (getIncludeHeadersFooters()) {
                        sheetExtractor.appendHeaderText(text);
                    }
                    sheetExtractor.appendCellText(text);
                    if (getIncludeTextBoxes()) {
                        processShapes(iter.getShapes(), text);
                    }
                    if (getIncludeHeadersFooters()) {
                        sheetExtractor.appendFooterText(text);
                    }
                    sheetExtractor.reset();
                }
            }

            return text.toString();
        } catch (IOException | OpenXML4JException | SAXException e) {
            Log.w(TAG, "Failed to load text", e);
            return "";
        }
    }

}
