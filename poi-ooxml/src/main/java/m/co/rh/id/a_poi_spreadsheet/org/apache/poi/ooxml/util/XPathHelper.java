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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.ooxml.util;

import android.util.Log;

import javax.xml.XMLConstants;
import javax.xml.namespace.QName;
import javax.xml.xpath.XPathFactory;

public final class XPathHelper {
    private static final String TAG = "XPathHelper";

    private static final String OSGI_ERROR =
            "Schemas (*.xsb) for <CLASS> can't be loaded - usually this happens when OSGI " +
                    "loading is used and the thread context classloader has no reference to " +
                    "the xmlbeans classes - please either verify if the <XSB>.xsb is on the " +
                    "classpath or alternatively try to use the poi-ooxml-full-x.x.jar";

    private static final String MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static final String MAC_DML_NS = "http://schemas.microsoft.com/office/mac/drawingml/2008/main";
    private static final QName ALTERNATE_CONTENT_TAG = new QName(MC_NS, "AlternateContent");
    // AlternateContentDocument.AlternateContent.type.getName();

    private XPathHelper() {
    }

    static final XPathFactory xpathFactory = XPathFactory.newInstance();

    static {
        trySetFeature(xpathFactory, XMLConstants.FEATURE_SECURE_PROCESSING, true);
    }

    public static XPathFactory getFactory() {
        return xpathFactory;
    }

    private static void trySetFeature(XPathFactory xpf, String feature, boolean enabled) {
        try {
            xpf.setFeature(feature, enabled);
        } catch (Exception e) {
            Log.w(TAG, String.format("XPathFactory Feature (%s) unsupported", feature), e);
        } catch (AbstractMethodError ame) {
            Log.w(TAG, String.format("Cannot set XPathFactory feature (%s) because outdated XML parser in classpath", feature), ame);
        }
    }


}
