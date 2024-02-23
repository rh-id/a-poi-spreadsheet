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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xddf.usermodel.text;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextSpacing;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextSpacingPoint;

@Beta
public class XDDFSpacingPoints extends XDDFSpacing {
    private CTTextSpacingPoint points;

    public XDDFSpacingPoints(double value) {
        this(CTTextSpacing.Factory.newInstance(), CTTextSpacingPoint.Factory.newInstance());
        if (spacing.isSetSpcPct()) {
            spacing.unsetSpcPct();
        }
        spacing.setSpcPts(points);
        setPoints(value);
    }

    @Internal
    protected XDDFSpacingPoints(CTTextSpacing parent, CTTextSpacingPoint points) {
        super(parent);
        this.points = points;
    }

    @Override
    public Kind getType() {
        return Kind.POINTS;
    }

    public double getPoints() {
        return points.getVal() * 0.01;
    }

    public void setPoints(double value) {
        points.setVal((int)(100 * value));
    }
}
