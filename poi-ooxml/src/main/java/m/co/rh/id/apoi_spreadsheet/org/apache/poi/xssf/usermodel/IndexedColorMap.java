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
package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel;

/**
 * Interface for color index to RGB mappings.
 * May be either the default, built-in mappings
 * or custom mappings defined in the document.
 */
public interface IndexedColorMap {

    /**
     * @param index color index to look up
     * @return the RGB array for the index, or null if the index is invalid/undefined
     */
    byte[] getRGB(int index);
}
