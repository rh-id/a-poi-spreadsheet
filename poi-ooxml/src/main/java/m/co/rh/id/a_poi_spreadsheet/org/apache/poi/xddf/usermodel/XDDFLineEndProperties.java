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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.xddf.usermodel;

import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.Internal;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineEndProperties;

@Beta
public class XDDFLineEndProperties {
    private CTLineEndProperties props;

    protected XDDFLineEndProperties(CTLineEndProperties properties) {
        this.props = properties;
    }

    @Internal
    protected CTLineEndProperties getXmlObject() {
        return props;
    }

    public LineEndLength getLength() {
        return LineEndLength.valueOf(props.getLen());
    }

    public void setLength(LineEndLength length) {
        props.setLen(length.underlying);
    }

    public LineEndType getType() {
        return LineEndType.valueOf(props.getType());
    }

    public void setType(LineEndType type) {
        props.setType(type.underlying);
    }

    public LineEndWidth getWidth() {
        return LineEndWidth.valueOf(props.getW());
    }

    public void setWidth(LineEndWidth width) {
        props.setW(width.underlying);
    }
}
