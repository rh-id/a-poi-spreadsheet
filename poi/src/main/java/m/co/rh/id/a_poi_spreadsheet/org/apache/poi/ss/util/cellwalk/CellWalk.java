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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.util.cellwalk;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ss.util.CellRangeAddress;

/**
 * Traverse cell range.
 */
public class CellWalk {

    private final Sheet sheet;
    private final CellRangeAddress range;
    private boolean traverseEmptyCells;


    public CellWalk(Sheet sheet, CellRangeAddress range) {
        this.sheet = sheet;
        this.range = range;
        this.traverseEmptyCells = false;
    }

    /**
     * Should we call handler on empty (blank) cells. Default is
     * false.
     *
     * @return true if handler should be called on empty (blank)
     *         cells, false otherwise.
     */
    public boolean isTraverseEmptyCells() {
        return traverseEmptyCells;
    }

    /**
     * Sets the traverseEmptyCells property.
     *
     * @param traverseEmptyCells new property value
     */
    public void setTraverseEmptyCells(boolean traverseEmptyCells) {
        this.traverseEmptyCells = traverseEmptyCells;
    }

    /**
     * Traverse cell range from top left to bottom right cell.
     *
     * @param handler handler to call on each appropriate cell
     */
    public void traverse(CellHandler handler) {
        int firstRow = range.getFirstRow();
        int lastRow = range.getLastRow();
        int firstColumn = range.getFirstColumn();
        int lastColumn = range.getLastColumn();
        final int width = lastColumn - firstColumn + 1;
        SimpleCellWalkContext ctx = new SimpleCellWalkContext();
        Row currentRow;
        Cell currentCell;

        for (ctx.rowNumber = firstRow; ctx.rowNumber <= lastRow; ++ctx.rowNumber) {
            currentRow = sheet.getRow(ctx.rowNumber);
            if (currentRow == null) {
                continue;
            }
            for (ctx.colNumber = firstColumn; ctx.colNumber <= lastColumn; ++ctx.colNumber) {
                currentCell = currentRow.getCell(ctx.colNumber);

                if (currentCell == null) {
                    continue;
                }
                if (isEmpty(currentCell) && !traverseEmptyCells) {
                    continue;
                }

                long rowSize = Math.multiplyExact(Math.subtractExact(ctx.rowNumber, firstRow), (long)width);

                ctx.ordinalNumber = Math.addExact(rowSize, (ctx.colNumber - firstColumn + 1));

                handler.onCell(currentCell, ctx);
            }
        }
    }

    private boolean isEmpty(Cell cell) {
        return (cell.getCellType() == CellType.BLANK);
    }

    /**
     * Inner class to hold walk context.
     */
    private static class SimpleCellWalkContext implements CellWalkContext {
        private long ordinalNumber;
        private int rowNumber;
        private int colNumber;

        @Override
        public long getOrdinalNumber() {
            return ordinalNumber;
        }

        @Override
        public int getRowNumber() {
            return rowNumber;
        }

        @Override
        public int getColumnNumber() {
            return colNumber;
        }
    }
}
