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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.binary;

import java.util.Objects;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.Internal;

/**
 * This is a read only record that maintains information about
 * a hyperlink.  In OOXML land, this information has to be merged
 * from 1) the sheet's .rels to get the url and 2) from after the
 * sheet data in they hyperlink section.
 *
 * The {@link #display} is often empty and should be filled from
 * the contents of the anchor cell.
 *
 * @since 3.16-beta3
 */
@Internal
public class XSSFHyperlinkRecord {

    private final CellRangeAddress cellRangeAddress;
    private final String relId;
    private String location;
    private String toolTip;
    private String display;

    XSSFHyperlinkRecord(CellRangeAddress cellRangeAddress, String relId, String location, String toolTip, String display) {
        this.cellRangeAddress = cellRangeAddress;
        this.relId = relId;
        this.location = location;
        this.toolTip = toolTip;
        this.display = display;
    }

    void setLocation(String location) {
        this.location = location;
    }

    void setToolTip(String toolTip) {
        this.toolTip = toolTip;
    }

    void setDisplay(String display) {
        this.display = display;
    }

    CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public String getRelId() {
        return relId;
    }

    public String getLocation() {
        return location;
    }

    public String getToolTip() {
        return toolTip;
    }

    public String getDisplay() {
        return display;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        XSSFHyperlinkRecord that = (XSSFHyperlinkRecord) o;

        return Objects.equals(cellRangeAddress, that.cellRangeAddress) &&
                Objects.equals(relId, that.relId) &&
                Objects.equals(location, that.location) &&
                Objects.equals(toolTip, that.toolTip) &&
                Objects.equals(display, that.display);
    }

    @Override
    public int hashCode() {
        return Objects.hash(cellRangeAddress, relId, location, toolTip, display);
    }

    @Override
    public String toString() {
        return "XSSFHyperlinkRecord{" +
                "cellRangeAddress=" + cellRangeAddress +
                ", relId='" + relId + '\'' +
                ", location='" + location + '\'' +
                ", toolTip='" + toolTip + '\'' +
                ", display='" + display + '\'' +
                '}';
    }
}
