
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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.ddf;

/**
 * This class stores the type and description of an escher property.
 */
public class EscherPropertyMetaData
{
    // Escher property types.
    public static final byte TYPE_UNKNOWN = (byte) 0;
    public static final byte TYPE_BOOLEAN = (byte) 1;
    public static final byte TYPE_RGB = (byte) 2;
    public static final byte TYPE_SHAPEPATH = (byte) 3;
    public static final byte TYPE_SIMPLE = (byte)4;
    public static final byte TYPE_ARRAY = (byte)5;

    private String description;
    private byte type;


    /**
     * @param description The description of the escher property.
     */
    public EscherPropertyMetaData( String description )
    {
        this.description = description;
    }

    /**
     *
     * @param description   The description of the escher property.
     * @param type          The type of the property.
     */
    public EscherPropertyMetaData( String description, byte type )
    {
        this.description = description;
        this.type = type;
    }

    public String getDescription()
    {
        return description;
    }

    public byte getType()
    {
        return type;
    }

}
