/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ConditionFilterData;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;

public class XSSFConditionFilterData implements ConditionFilterData {

    private final CTCfRule _cfRule;
    
    /*package*/ XSSFConditionFilterData(CTCfRule cfRule) {
        _cfRule = cfRule;
    }

    public boolean getAboveAverage() {
        return _cfRule.getAboveAverage();
    }

    public boolean getBottom() {
        return _cfRule.getBottom();
    }

    public boolean getEqualAverage() {
        return _cfRule.getEqualAverage();
    }

    public boolean getPercent() {
        return _cfRule.getPercent();
    }

    public long getRank() {
        return _cfRule.getRank();
    }

    public int getStdDev() {
        return _cfRule.getStdDev();
    }

}
