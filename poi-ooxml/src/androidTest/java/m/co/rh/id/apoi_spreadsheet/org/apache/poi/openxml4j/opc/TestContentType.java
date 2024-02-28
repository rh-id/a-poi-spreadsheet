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

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertThrows;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import android.util.Log;

import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;

import m.co.rh.id.apoi_spreadsheet.POIJUnit4ClassRunner;
import m.co.rh.id.apoi_spreadsheet.POIJUnit4Parameterized;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.OpenXML4JTestDataSamples;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.openxml4j.opc.internal.ContentType;

/**
 * Tests for content type (ContentType class).
 */
@RunWith(Enclosed.class)
public final class TestContentType {

    private static final String FEATURE_DISALLOW_DOCTYPE_DECL = "http://apache.org/xml/features/disallow-doctype-decl";

    @RunWith(POIJUnit4Parameterized.class)
    public static class ContentTypeValidation {
        @Parameterized.Parameters(name = "{index}: testContentTypeValidation({0})")
        public static Iterable<Object[]> data() {
            List<Object[]> data = new ArrayList<>();
            data.add(new Object[]{"text/xml"});
            data.add(new Object[]{"application/pgp-key"});
            data.add(new Object[]{"application/vnd.hp-PCLXL"});
            data.add(new Object[]{"application/vnd.lotus-1-2-3"});
            return data;
        }

        @Parameterized.Parameter
        public String contentType;

        /**
         * Check rule M1.13: Package implementers shall only create and only
         * recognize parts with a content type; format designers shall specify a
         * content type for each part included in the format. Content types for
         * package parts shall fit the definition and syntax for media types as
         * specified in RFC 2616, \u00A73.7.
         */
        @Test
        public void testContentTypeValidation() throws InvalidFormatException {
            new ContentType(contentType);
            //assertDoesNotThrow(() -> new ContentType(contentType));
        }
    }


    @RunWith(POIJUnit4Parameterized.class)
    public static class ContentTypeValidationFailure {
        @Parameterized.Parameters(name = "{index}: testContentTypeValidationFailure({0})")
        public static Iterable<Object[]> data() {
            List<Object[]> data = new ArrayList<>();
            data.add(new Object[]{"text/xml/app"});
            data.add(new Object[]{""});
            data.add(new Object[]{"test"});
            data.add(new Object[]{"text(xml/xml"});
            data.add(new Object[]{"text)xml/xml"});
            data.add(new Object[]{"text<xml/xml"});
            data.add(new Object[]{"text>/xml"});
            data.add(new Object[]{"text@/xml"});
            data.add(new Object[]{"text,/xml"});
            data.add(new Object[]{"text;/xml"});
            data.add(new Object[]{"text:/xml"});
            data.add(new Object[]{"text\\/xml"});
            data.add(new Object[]{"t/ext/xml"});
            data.add(new Object[]{"t\"ext/xml"});
            data.add(new Object[]{"text[/xml"});
            data.add(new Object[]{"text]/xml"});
            data.add(new Object[]{"text?/xml"});
            data.add(new Object[]{"tex=t/xml"});
            data.add(new Object[]{"te{xt/xml"});
            data.add(new Object[]{"tex}t/xml"});
            data.add(new Object[]{"te xt/xml"});
            data.add(new Object[]{"text\u0009/xml"});
            data.add(new Object[]{"text xml"});
            data.add(new Object[]{"text/xml "});
            return data;
        }

        @Parameterized.Parameter
        public String contentType;

        /**
         * Check rule M1.13 : Package implementers shall only create and only
         * recognize parts with a content type; format designers shall specify a
         * content type for each part included in the format. Content types for
         * package parts shall fit the definition and syntax for media types as
         * specified in RFC 2616, \u00A3.7.
         * <p>
         * Check rule M1.14: Content types shall not use linear white space either
         * between the type and subtype or between an attribute and its value.
         * Content types also shall not have leading or trailing white spaces.
         * Package implementers shall create only such content types and shall
         * require such content types when retrieving a part from a package; format
         * designers shall specify only such content types for inclusion in the
         * format.
         */
        @Test
        public void testContentTypeValidationFailure() {
            assertThrows("Must have fail for content type: '" + contentType + "' !",
                    InvalidFormatException.class, () -> new ContentType(contentType));
        }

    }

    @RunWith(POIJUnit4Parameterized.class)
    public static class ContentTypeParam {
        @Parameterized.Parameters(name = "{index}: testContentTypeParam({0})")
        public static Iterable<Object[]> data() {
            List<Object[]> data = new ArrayList<>();
            data.add(new Object[]{"mail/toto;titi=tata"});
            data.add(new Object[]{"text/xml;a=b;c=d"});
            data.add(new Object[]{"text/xml;key1=param1;key2=param2"});
            data.add(new Object[]{"application/pgp-key;version=\"2\""});
            data.add(new Object[]{"application/x-resqml+xml;version=2.0;type=obj_global2dCrs"});
            return data;
        }

        @Parameterized.Parameter
        public String contentType;

        /**
         * Parameters are allowed, provides that they meet the
         * criteria of rule [01.2]
         * Invalid parameters are verified as incorrect in
         */
        @Test
        public void testContentTypeParam() throws InvalidFormatException {
            new ContentType(contentType);
            //assertDoesNotThrow(() -> new ContentType(contentType));
        }
    }

    @RunWith(POIJUnit4Parameterized.class)
    public static class ContentTypeParameterFailure {
        @Parameterized.Parameters(name = "{index}: testContentTypeParameterFailure({0})")
        public static Iterable<Object[]> data() {
            List<Object[]> data = new ArrayList<>();
            data.add(new Object[]{"mail/toto;\"titi=tata\""});
            data.add(new Object[]{"mail/toto;titi = tata"});
            data.add(new Object[]{"text/\u0080"});
            return data;
        }

        @Parameterized.Parameter
        public String contentType;

        /**
         * Check rule [O1.2]: Format designers might restrict the usage of
         * parameters for content types.
         */
        @Test
        public void testContentTypeParameterFailure() {
            assertThrows("Must have fail for content type: '" + contentType + "' !",
                    InvalidFormatException.class, () -> new ContentType(contentType));
        }
    }

    @RunWith(POIJUnit4Parameterized.class)
    public static class ContentTypeCommentFailure {
        @Parameterized.Parameters(name = "{index}: testContentTypeCommentFailure({0})")
        public static Iterable<Object[]> data() {
            List<Object[]> data = new ArrayList<>();
            data.add(new Object[]{"text/xml(comment)"});
            return data;
        }

        @Parameterized.Parameter
        public String contentType;

        /**
         * Check rule M1.15: The package implementer shall require a content type
         * that does not include comments and the format designer shall specify such
         * a content type.
         */
        @Test
        public void testContentTypeCommentFailure() {
            assertThrows("Must have fail for content type: '" + contentType + "' !",
                    InvalidFormatException.class, () -> new ContentType(contentType));
        }
    }

    @RunWith(POIJUnit4ClassRunner.class)
    public static class ContentTypeOtherTests{
        /**
         * OOXML content types don't need entities and we shouldn't
         * barf if we get one from a third party system that added them
         * (expected = InvalidFormatException.class)
         */
        @Test
        public void testFileWithContentTypeEntities() throws Exception {
            try (InputStream is = OpenXML4JTestDataSamples.openSampleStream("ContentTypeHasEntities.ooxml")) {
                /* TODO why isOldXercesActive logic like this?
                if (isOldXercesActive()) {
                    OPCPackage.open(is);
                } else {
                    assertThrows(InvalidFormatException.class, () -> OPCPackage.open(is));
                }*/
                assertThrows(InvalidFormatException.class, () -> OPCPackage.open(is));
            }
        }

        /**
         * Check that we can open a file where there are valid
         * parameters on a content type
         */
        @Test
        public void testFileWithContentTypeParams() throws Exception {
            try (InputStream is = OpenXML4JTestDataSamples.openSampleStream("ContentTypeHasParameters.ooxml");
                 OPCPackage p = OPCPackage.open(is)) {

                final String typeResqml = "application/x-resqml+xml";

                // Check the types on everything
                for (PackagePart part : p.getParts()) {
                    final String contentType = part.getContentType();
                    final ContentType details = part.getContentTypeDetails();
                    final int length = details.getParameterKeys().length;
                    final boolean hasParameters = details.hasParameters();

                    // _rels type doesn't have any params
                    if (part.isRelationshipPart()) {
                        assertEquals(ContentTypes.RELATIONSHIPS_PART, contentType);
                        assertEquals(ContentTypes.RELATIONSHIPS_PART, details.toString());
                        assertFalse(hasParameters);
                        assertEquals(0, length);
                    }
                    // Core type doesn't have any params
                    else if (part.getPartName().toString().equals("/docProps/core.xml")) {
                        assertEquals(ContentTypes.CORE_PROPERTIES_PART, contentType);
                        assertEquals(ContentTypes.CORE_PROPERTIES_PART, details.toString());
                        assertFalse(hasParameters);
                        assertEquals(0, length);
                    }
                    // Global Crs types do have params
                    else if (part.getPartName().toString().equals("/global1dCrs.xml")) {
                        assertTrue(part.getContentType().startsWith(typeResqml));
                        assertEquals(typeResqml, details.toString(false));
                        assertTrue(hasParameters);
                        assertContains("version=2.0", details.toString());
                        assertContains("type=obj_global1dCrs", details.toString());
                        assertEquals(2, length);
                        assertEquals("2.0", details.getParameter("version"));
                        assertEquals("obj_global1dCrs", details.getParameter("type"));
                    } else if (part.getPartName().toString().equals("/global2dCrs.xml")) {
                        assertTrue(part.getContentType().startsWith(typeResqml));
                        assertEquals(typeResqml, details.toString(false));
                        assertTrue(hasParameters);
                        assertContains("version=2.0", details.toString());
                        assertContains("type=obj_global2dCrs", details.toString());
                        assertEquals(2, length);
                        assertEquals("2.0", details.getParameter("version"));
                        assertEquals("obj_global2dCrs", details.getParameter("type"));
                    }
                    // Other thingy
                    else if (part.getPartName().toString().equals("/myTestingGuid.xml")) {
                        assertTrue(part.getContentType().startsWith(typeResqml));
                        assertEquals(typeResqml, details.toString(false));
                        assertTrue(hasParameters);
                        assertContains("version=2.0", details.toString());
                        assertContains("type=obj_tectonicBoundaryFeature", details.toString());
                        assertEquals(2, length);
                        assertEquals("2.0", details.getParameter("version"));
                        assertEquals("obj_tectonicBoundaryFeature", details.getParameter("type"));
                    }
                    // That should be it!
                    else {
                        fail("Unexpected part " + part);
                    }
                }
            }
        }
    }

    private static void assertContains(String needle, String haystack) {
        assertTrue(haystack.contains(needle));
    }

    public static boolean isOldXercesActive() {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        try {
            dbf.setFeature(FEATURE_DISALLOW_DOCTYPE_DECL, true);
            return false;
        } catch (Exception | AbstractMethodError ignored) {
            Log.e("isOldXercesActive", ignored.getMessage(), ignored);
        }
        return true;
    }
}
