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

package m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util;

import android.graphics.Bitmap;
import android.graphics.Color;
import android.graphics.Matrix;
import android.graphics.Path;
import android.graphics.Point;
import android.graphics.PointF;
import android.graphics.Rect;

import androidx.graphics.path.PathIterator;
import androidx.graphics.path.PathSegment;
import androidx.graphics.path.PathUtilities;

import org.apache.commons.io.output.NullOutputStream;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.lang.reflect.Array;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import m.co.rh.id.a_poi_spreadsheet.base.model.Dimension2D;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.common.usermodel.GenericRecord;
import m.co.rh.id.a_poi_spreadsheet.org.apache.poi.util.GenericRecordJsonWriter.AppendableWriter;

@SuppressWarnings("WeakerAccess")
public class GenericRecordXmlWriter implements Closeable {
    private static final String TABS;
    private static final String ZEROS = "0000000000000000";
    private static final Pattern ESC_CHARS = Pattern.compile("[<>&'\"\\p{Cntrl}]");

    @FunctionalInterface
    protected interface GenericRecordHandler {
        /**
         * Handler method
         *
         * @param record the parent record, applied via instance method reference
         * @param name   the name of the property
         * @param object the value of the property
         * @return {@code true}, if the element was handled and output produced,
         * The provided methods can be overridden and an implementation can return {@code false},
         * if the element hasn't been written to the stream
         */
        boolean print(GenericRecordXmlWriter record, String name, Object object);
    }

    private static final List<Map.Entry<Class<?>, GenericRecordHandler>> handler = new ArrayList<>();

    static {
        char[] t = new char[255];
        Arrays.fill(t, '\t');
        TABS = new String(t);
        handler(String.class, GenericRecordXmlWriter::printObject);
        handler(Number.class, GenericRecordXmlWriter::printNumber);
        handler(Boolean.class, GenericRecordXmlWriter::printBoolean);
        handler(List.class, GenericRecordXmlWriter::printList);
        // handler(GenericRecord.class, GenericRecordXmlWriter::printGenericRecord);
        handler(GenericRecordUtil.AnnotatedFlag.class, GenericRecordXmlWriter::printAnnotatedFlag);
        handler(byte[].class, GenericRecordXmlWriter::printBytes);
        handler(Point.class, GenericRecordXmlWriter::printPoint);
        handler(Dimension2D.class, GenericRecordXmlWriter::printDimension);
        handler(Rect.class, GenericRecordXmlWriter::printRectangle);
        handler(Path.class, GenericRecordXmlWriter::printPath);
        handler(Matrix.class, GenericRecordXmlWriter::printAffineTransform);
        handler(Color.class, GenericRecordXmlWriter::printColor);
        handler(Bitmap.class, GenericRecordXmlWriter::printBufferedImage);
        handler(Array.class, GenericRecordXmlWriter::printArray);
        handler(Object.class, GenericRecordXmlWriter::printObject);
    }

    private static void handler(Class<?> c, GenericRecordHandler printer) {
        handler.add(new AbstractMap.SimpleEntry<>(c, printer));
    }

    private final PrintWriter fw;
    private int indent = 0;
    private boolean withComments = true;
    private int childIndex = 0;
    private boolean attributePhase = true;

    public GenericRecordXmlWriter(File fileName) throws IOException {
        OutputStream os = ("null".equals(fileName.getName())) ?
                NullOutputStream.INSTANCE : Files.newOutputStream(fileName.toPath());
        fw = new PrintWriter(new OutputStreamWriter(os, StandardCharsets.UTF_8));
    }

    public GenericRecordXmlWriter(Appendable buffer) {
        fw = new PrintWriter(new AppendableWriter(buffer));
    }

    public static String marshal(GenericRecord record) {
        return marshal(record, true);
    }

    public static String marshal(GenericRecord record, boolean withComments) {
        final StringBuilder sb = new StringBuilder();
        try (GenericRecordXmlWriter w = new GenericRecordXmlWriter(sb)) {
            w.setWithComments(withComments);
            w.write(record);
            return sb.toString();
        } catch (IOException e) {
            return "<record/>";
        }
    }

    public void setWithComments(boolean withComments) {
        this.withComments = withComments;
    }

    @Override
    public void close() throws IOException {
        fw.close();
    }

    protected String tabs() {
        return TABS.substring(0, Math.min(indent, TABS.length()));
    }

    public void write(GenericRecord record) {
        write("record", record);
    }

    protected void write(final String name, GenericRecord record) {
        final String tabs = tabs();
        Enum<?> type = record.getGenericRecordType();
        String recordName = (type != null) ? type.name() : record.getClass().getSimpleName();
        fw.append(tabs);
        fw.append("<").append(name).append(" type=\"");
        fw.append(recordName);
        fw.append("\"");
        if (childIndex > 0) {
            fw.append(" index=\"");
            fw.print(childIndex);
            fw.append("\"");
        }


        attributePhase = true;

        boolean hasComplex = writeProperties(record);

        attributePhase = false;

        hasComplex |= writeChildren(record, hasComplex);

        if (hasComplex) {
            fw.append(tabs);
            fw.println("</" + name + ">");
        } else {
            fw.println("/>");
        }
    }

    protected boolean writeProperties(GenericRecord record) {
        Map<String, Supplier<?>> prop = record.getGenericProperties();
        if (prop == null || prop.isEmpty()) {
            return false;
        }

        final int oldChildIndex = childIndex;
        childIndex = 0;
        List<Map.Entry<String, Supplier<?>>> complex = prop.entrySet().stream().flatMap(this::writeProp).collect(Collectors.toList());

        attributePhase = false;
        if (!complex.isEmpty()) {
            fw.println(">");
            indent++;
            complex.forEach(this::writeProp);
            indent--;
        }
        childIndex = oldChildIndex;

        return !complex.isEmpty();
    }

    protected boolean writeChildren(GenericRecord record, boolean hasComplexProperties) {
        List<? extends GenericRecord> list = record.getGenericChildren();
        if (list == null || list.isEmpty()) {
            return false;
        }
        if (!hasComplexProperties) {
            fw.print(">");
        }

        indent++;
        fw.println();
        fw.println(tabs() + "<children>");
        indent++;
        final int oldChildIndex = childIndex;
        childIndex = 0;
        list.forEach(l -> {
            writeValue("record", l);
            childIndex++;
        });
        childIndex = oldChildIndex;
        fw.println();
        indent--;
        fw.println(tabs() + "</children>");
        indent--;

        return true;
    }

    public void writeError(String errorMsg) {
        printObject("error", errorMsg);
    }

    @SuppressWarnings("java:S1452")
    protected Stream<Map.Entry<String, Supplier<?>>> writeProp(Map.Entry<String, Supplier<?>> me) {
        Object obj = me.getValue().get();
        if (obj == null) {
            return Stream.empty();
        }

        final boolean isComplex = isComplex(obj);
        if (attributePhase == isComplex) {
            return isComplex ? Stream.of(new AbstractMap.SimpleEntry<>(me.getKey(), () -> obj)) : Stream.empty();
        }

        final int oldChildIndex = childIndex;
        childIndex = 0;
        writeValue(me.getKey(), obj);
        childIndex = oldChildIndex;

        return Stream.empty();
    }

    protected static boolean isComplex(Object obj) {
        return !(
                obj instanceof Number ||
                        obj instanceof Boolean ||
                        obj instanceof Character ||
                        obj instanceof String ||
                        obj instanceof Color ||
                        obj instanceof Enum);
    }

    protected void writeValue(String name, Object value) {
        assert (name != null);
        if (value instanceof GenericRecord) {
            printGenericRecord(name, value);
        } else if (value != null) {
            if (name.endsWith(">")) {
                fw.print("\t");
            }

            handler.stream().
                    filter(h -> matchInstanceOrArray(h.getKey(), value)).
                    findFirst().
                    ifPresent(h -> h.getValue().print(this, name, value));

        }
    }

    protected static boolean matchInstanceOrArray(Class<?> key, Object instance) {
        return key.isInstance(instance) || (Array.class.equals(key) && instance.getClass().isArray());
    }

    protected void openName(String name) {
        name = name.replace(">>", ">");
        if (attributePhase) {
            fw.print(" " + name.replace('>', ' ').trim() + "=\"");
        } else {
            fw.print(tabs() + "<" + name);
            if (name.endsWith(">")) {
                fw.println();
            }
        }
    }

    protected void closeName(String name) {
        name = name.replace(">>", ">");
        if (attributePhase) {
            fw.append("\"");
        } else {
            if (name.endsWith(">")) {
                fw.println(tabs() + "\t</" + name);
            } else {
                fw.println("/>");
            }
        }
    }

    protected boolean printNumber(String name, Object o) {
        assert (attributePhase);

        openName(name);
        Number n = (Number) o;
        fw.print(n.toString());
        closeName(name);

        return true;
    }

    protected boolean printBoolean(String name, Object o) {
        assert (attributePhase);
        openName(name);
        fw.write(((Boolean) o).toString());
        closeName(name);
        return true;
    }

    protected boolean printList(String name, Object o) {
        assert (!attributePhase);
        openName(name + ">");
        int oldChildIndex = childIndex;
        childIndex = 0;
        ((List<?>) o).forEach(e -> {
            writeValue("item>", e);
            childIndex++;
        });
        childIndex = oldChildIndex;
        closeName(name + ">");
        return true;
    }

    protected boolean printArray(String name, Object o) {
        assert (!attributePhase);
        openName(name + ">");
        int length = Array.getLength(o);
        final int oldChildIndex = childIndex;
        for (childIndex = 0; childIndex < length; childIndex++) {
            writeValue("item>", Array.get(o, childIndex));
        }
        childIndex = oldChildIndex;
        closeName(name + ">");
        return true;
    }

    protected void printGenericRecord(String name, Object value) {
        write(name, (GenericRecord) value);
    }

    protected boolean printAnnotatedFlag(String name, Object o) {
        assert (!attributePhase);
        GenericRecordUtil.AnnotatedFlag af = (GenericRecordUtil.AnnotatedFlag) o;
        Number n = af.getValue().get();
        int len;
        if (n instanceof Byte) {
            len = 2;
        } else if (n instanceof Short) {
            len = 4;
        } else if (n instanceof Integer) {
            len = 8;
        } else {
            len = 16;
        }

        openName(name);
        fw.print(" flag=\"0x");
        fw.print(trimHex(n.longValue(), len));
        fw.print('"');
        if (withComments) {
            fw.print(" description=\"");
            fw.print(af.getDescription());
            fw.print("\"");
        }
        closeName(name);
        return true;
    }

    protected boolean printBytes(String name, Object o) {
        assert (!attributePhase);
        openName(name + ">");
        fw.write(Base64.getEncoder().encodeToString((byte[]) o));
        closeName(name + ">");
        return true;
    }

    protected boolean printPoint(String name, Object o) {
        assert (!attributePhase);
        openName(name);
        Point p = (Point) o;
        fw.println(" x=\"" + p.x + "\" y=\"" + p.y + "\"/>");
        closeName(name);
        return true;
    }

    protected boolean printDimension(String name, Object o) {
        assert (!attributePhase);
        openName(name);
        Dimension2D p = (Dimension2D) o;
        fw.println(" width=\"" + p.getWidth() + "\" height=\"" + p.getHeight() + "\"/>");
        closeName(name);
        return true;
    }

    protected boolean printRectangle(String name, Object o) {
        assert (!attributePhase);
        openName(name);
        Rect p = (Rect) o;
        fw.println(" x=\"" + p.left + "\" y=\"" + p.top + "\" width=\"" + p.width() + "\" height=\"" + p.height() + "\"/>");
        closeName(name);
        return true;
    }

    protected boolean printPath(String name, Object o) {
        assert (!attributePhase);

        openName(name + ">");

        final PathIterator iter = PathUtilities.iterator((Path) o);

        indent += 2;
        String t = tabs();
        indent -= 2;

        while (!iter.hasNext()) {
            fw.print(t);
            PathSegment pathSegment = iter.next();
            final PathSegment.Type segType = pathSegment.getType();
            fw.print("<pathelement ");
            switch (segType) {
                case Move:
                    PointF movePts = pathSegment.getPoints()[0];
                    fw.print("type=\"move\" x=\"" + movePts.x + "\" y=\"" + movePts.y + "\"");
                    break;
                case Line:
                    PointF linePts = pathSegment.getPoints()[1];
                    fw.print("type=\"lineto\" x=\"" + linePts.x + "\" y=\"" + linePts.y + "\"");
                    break;
                case Quadratic:
                    PointF qControlPts = pathSegment.getPoints()[1];
                    PointF qEndPts = pathSegment.getPoints()[2];
                    fw.print("type=\"quad\" x1=\"" + qControlPts.x + "\" y1=\"" + qControlPts.y + "\" x2=\"" + qEndPts.x + "\" y2=\"" + qEndPts.y + "\"");
                    break;
                case Cubic:
                    PointF cFirstCtrlPts = pathSegment.getPoints()[1];
                    PointF cSecondCtrlPts = pathSegment.getPoints()[2];
                    PointF cEndCtrlPts = pathSegment.getPoints()[3];
                    fw.print("type=\"cubic\" x1=\"" + cFirstCtrlPts.x + "\" y1=\"" + cFirstCtrlPts.y + "\" x2=\"" + cSecondCtrlPts.x + "\" y2=\"" + cSecondCtrlPts.y + "\" x3=\"" + cEndCtrlPts.x + "\" y3=\"" + cEndCtrlPts.y + "\"");
                    break;
                case Close:
                    fw.print("type=\"close\"");
                    break;
            }
            fw.println("/>");
            iter.next();
        }

        closeName(name + ">");
        return true;
    }

    protected boolean printObject(String name, Object o) {
        openName(name + ">");
        final String str = o.toString();
        final Matcher m = ESC_CHARS.matcher(str);
        int pos = 0;
        while (m.find()) {
            fw.write(str, pos, m.start());
            String match = m.group();
            switch (match) {
                case "<":
                    fw.write("&lt;");
                    break;
                case ">":
                    fw.write("&gt;");
                    break;
                case "&":
                    fw.write("&amp;");
                    break;
                case "'":
                    fw.write("&apos;");
                    break;
                case "\"":
                    fw.write("&quot;");
                    break;
                default:
                    fw.write("&#x");
                    fw.write(Long.toHexString(match.codePointAt(0)));
                    fw.write(";");
                    break;
            }
            pos = m.end();
        }
        fw.append(str, pos, str.length());
        closeName(name + ">");
        return true;
    }

    protected boolean printAffineTransform(String name, Object o) {
        assert (!attributePhase);
        openName(name);
        Matrix xForm = (Matrix) o;
        float[] values = new float[9];
        xForm.getValues(values);
        fw.write("<" + name +
                " scaleX=\"" + values[0] + "\" " +
                "shearX=\"" + values[1] + "\" " +
                "transX=\"" + values[2] + "\" " +
                "scaleY=\"" + values[4] + "\" " +
                "shearY=\"" + values[3] + "\" " +
                "transY=\"" + values[5] + "\"/>");
        closeName(name);
        return true;
    }

    protected boolean printColor(String name, Object o) {
        assert (attributePhase);
        openName(name);
        final int rgb = ((Color) o).toArgb();
        fw.print("0x" + trimHex(rgb, 8));
        closeName(name);
        return true;
    }

    protected boolean printBufferedImage(String name, Object o) {
        assert (!attributePhase);
        openName(name);
        Bitmap bi = (Bitmap) o;
        fw.println(" width=\"" + bi.getWidth() + "\" height=\"" + bi.getHeight() + "\"");
        closeName(name);
        return true;
    }

    protected String trimHex(final long l, final int size) {
        final String b = Long.toHexString(l);
        int len = b.length();
        return ZEROS.substring(0, Math.max(0, size - len)) + b.substring(Math.max(0, len - size), len);
    }

}