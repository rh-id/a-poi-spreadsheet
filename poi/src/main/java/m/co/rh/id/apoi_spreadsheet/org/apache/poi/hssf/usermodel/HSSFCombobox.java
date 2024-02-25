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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherBoolProperty;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherClientDataRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherContainerRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherOptRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherPropertyTypes;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherSimpleProperty;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf.EscherSpRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.CommonObjectDataSubRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.EndSubRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.FtCblsSubRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.LbsDataSubRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.ObjRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.record.TextObjectRecord;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.sl.usermodel.ShapeType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;

/**
 *
 */
public class HSSFCombobox extends HSSFSimpleShape {

    public HSSFCombobox(EscherContainerRecord spContainer, ObjRecord objRecord) {
        super(spContainer, objRecord);
    }

    public HSSFCombobox(HSSFShape parent, HSSFAnchor anchor) {
        super(parent, anchor);
        super.setShapeType(OBJECT_TYPE_COMBO_BOX);
        CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord) getObjRecord().getSubRecords().get(0);
        cod.setObjectType(CommonObjectDataSubRecord.OBJECT_TYPE_COMBO_BOX);
    }

    @Override
    protected TextObjectRecord createTextObjRecord() {
        return null;
    }

    @Override
    protected EscherContainerRecord createSpContainer() {
        EscherContainerRecord spContainer = new EscherContainerRecord();
        EscherSpRecord sp = new EscherSpRecord();
        EscherOptRecord opt = new EscherOptRecord();
        EscherClientDataRecord clientData = new EscherClientDataRecord();

        spContainer.setRecordId(EscherContainerRecord.SP_CONTAINER);
        spContainer.setOptions((short) 0x000F);
        sp.setRecordId(EscherSpRecord.RECORD_ID);
        sp.setOptions((short) ((ShapeType.HOST_CONTROL.nativeId << 4) | 0x2));

        sp.setFlags(EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
        opt.setRecordId(EscherOptRecord.RECORD_ID);
        opt.addEscherProperty(new EscherBoolProperty(EscherPropertyTypes.PROTECTION__LOCKAGAINSTGROUPING, 17039620));
        opt.addEscherProperty(new EscherBoolProperty(EscherPropertyTypes.TEXT__SIZE_TEXT_TO_FIT_SHAPE, 0x00080008));
        opt.addEscherProperty(new EscherBoolProperty(EscherPropertyTypes.LINESTYLE__NOLINEDRAWDASH, 0x00080000));
        opt.addEscherProperty(new EscherSimpleProperty(EscherPropertyTypes.GROUPSHAPE__FLAGS, 0x00020000));

        HSSFClientAnchor userAnchor = (HSSFClientAnchor) getAnchor();
        userAnchor.setAnchorType(AnchorType.DONT_MOVE_DO_RESIZE);
        EscherRecord anchor = userAnchor.getEscherAnchor();
        clientData.setRecordId(EscherClientDataRecord.RECORD_ID);
        clientData.setOptions((short) 0x0000);

        spContainer.addChildRecord(sp);
        spContainer.addChildRecord(opt);
        spContainer.addChildRecord(anchor);
        spContainer.addChildRecord(clientData);

        return spContainer;
    }

    @Override
    protected ObjRecord createObjRecord() {
        ObjRecord obj = new ObjRecord();
        CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
        c.setObjectType(HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX);
        c.setLocked(true);
        c.setPrintable(false);
        c.setAutofill(true);
        c.setAutoline(false);
        FtCblsSubRecord f = new FtCblsSubRecord();
        LbsDataSubRecord l = LbsDataSubRecord.newAutoFilterInstance();
        EndSubRecord e = new EndSubRecord();
        obj.addSubRecord(c);
        obj.addSubRecord(f);
        obj.addSubRecord(l);
        obj.addSubRecord(e);
        return obj;
    }

    @Override
    public void setShapeType(int shapeType) {
        throw new IllegalStateException("Shape type can not be changed in " + this.getClass().getSimpleName());
    }
}
