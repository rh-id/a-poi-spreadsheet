package m.co.rh.id.apoi_spreadsheet.r8smoke;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import android.content.Context;
import android.util.Log;

import androidx.test.ext.junit.runners.AndroidJUnit4;
import androidx.test.platform.app.InstrumentationRegistry;

import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.TimeZone;

import javax.xml.stream.XMLInputFactory;

import m.co.rh.id.apoi_spreadsheet.base.POISpreadsheetContext;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.ExtractorFactory;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Cell;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellStyle;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.CellType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FillPatternType;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.DataFormatter;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Font;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.FormulaEvaluator;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.IndexedColors;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Row;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Sheet;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.Workbook;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.WorkbookFactory;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Permanent consumer-side smoke harness: these tests execute against the MINIFIED release APK
 * (testBuildType 'release') to empirically guard the consumer ProGuard/R8 rules that ship inside
 * the published AAR ({@code a-poi-spreadsheet/consumer-rules.pro}, merged from the consumer
 * rules of the former poi/poi-ooxml modules).
 *
 * <p>A failure here means a consumer-rule regression: R8 renamed/removed a class that POI loads
 * reflectively (aalto StAX implementations, XMLBeans schema types, ServiceLoader factories,
 * XML-DSig provider) or broke a signature the library resolves by name. Fix the rules in the
 * library's consumer-rules.pro file - not in this harness - then re-run.</p>
 */
@RunWith(AndroidJUnit4.class)
public class R8SmokeTest {

    private static final String TAG = "R8SmokeTest";

    private Context context;

    @Before
    public void setUp() {
        TimeZone.setDefault(TimeZone.getTimeZone("UTC"));
        context = InstrumentationRegistry.getInstrumentation().getTargetContext();
        POISpreadsheetContext.getInstance().setAppContext(context);
    }

    private File newCacheFile(String name) {
        return new File(context.getCacheDir(), name);
    }

    /** Builds a.xlsx under cacheDir; layout: A1='Hello'(blue fill+bold), B1=123.45, C1=date, A2=10, B2=32, C2=SUM(A2,B2). */
    private File buildXlsx() throws Exception {
        File file = newCacheFile("a.xlsx");
        Date dateValue = new GregorianCalendar(2026, Calendar.AUGUST, 28, 12, 0, 0).getTime();
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("S1");

            // Row 0 -> Excel refs A1..C1
            Row row0 = sheet.createRow(0);
            Cell stringCell = row0.createCell(0);
            stringCell.setCellValue("Hello");

            Cell numericCell = row0.createCell(1);
            numericCell.setCellValue(123.45);

            Cell dateCell = row0.createCell(2);
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
            dateCell.setCellStyle(dateStyle);
            dateCell.setCellValue(dateValue);

            // Row 1 -> Excel refs A2..C2
            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(10);
            row1.createCell(1).setCellValue(32);
            row1.createCell(2).setCellFormula("SUM(A2,B2)");

            // Cell fill (IndexedColors.BLUE foreground) + bold font on the string cell
            CellStyle styled = wb.createCellStyle();
            Font bold = wb.createFont();
            bold.setBold(true);
            styled.setFont(bold);
            styled.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            styled.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            stringCell.setCellStyle(styled);

            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
        }
        return file;
    }

    /** Builds h.xls under cacheDir; layout: A1=20, B1=22, C1=SUM(A1,B1), D1='Hello'. */
    private File buildHssf() throws Exception {
        File file = newCacheFile("h.xls");
        try (Workbook wb = new HSSFWorkbook()) {
            Sheet sheet = wb.createSheet("H1");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue(20);
            row0.createCell(1).setCellValue(22);
            row0.createCell(2).setCellFormula("SUM(A1,B1)");
            row0.createCell(3).setCellValue("Hello");

            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
        }
        return file;
    }

    /** T1 - XSSF (.xlsx) round-trip through WorkbookFactory (full parse path). */
    @Test
    public void xssfRoundTrip() throws Exception {
        File file = buildXlsx();

        // Reopen via WorkbookFactory -> ServiceLoader XSSFWorkbookFactory + OPCPackage
        // + XMLBeans TypeSystemHolder + reflective schema impls + aalto StAX
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = wb.getSheet("S1");
            assertNotNull("Sheet S1 missing after round-trip", sheet);

            Cell stringCell = sheet.getRow(0).getCell(0);
            assertEquals("Hello", stringCell.getStringCellValue());

            assertEquals(123.45, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.0001);

            Cell formulaCell = sheet.getRow(1).getCell(2);
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(formulaCell);
            assertEquals("SUM formula must evaluate to 42", 42.0, formulaCell.getNumericCellValue(), 0.0001);

            String formattedDate = new DataFormatter()
                    .formatCellValue(sheet.getRow(0).getCell(2));
            Log.i(TAG, "formatted date cell: " + formattedDate);
            assertNotNull(formattedDate);
            assertTrue("date cell must format to non-empty text", !formattedDate.trim().isEmpty());

            CellStyle reloaded = stringCell.getCellStyle();
            assertEquals("fill foreground color must survive",
                    IndexedColors.BLUE.getIndex(), reloaded.getFillForegroundColor());
            assertEquals("fill pattern must survive",
                    FillPatternType.SOLID_FOREGROUND, reloaded.getFillPattern());
            assertTrue("bold font must survive",
                    wb.getFontAt(reloaded.getFontIndexAsInt()).getBold());
        }
    }

    /** T2 - HSSF (.xls) round-trip through WorkbookFactory (ServiceLoader + POIFS). */
    @Test
    public void hssfRoundTrip() throws Exception {
        File file = buildHssf();

        try (Workbook wb = WorkbookFactory.create(file)) {
            Cell stringCell = wb.getSheet("H1").getRow(0).getCell(3);
            assertEquals("Hello", stringCell.getStringCellValue());

            Cell formulaCell = wb.getSheet("H1").getRow(0).getCell(2);
            assertEquals(CellType.FORMULA, formulaCell.getCellType());
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(formulaCell);
            assertEquals("SUM formula must evaluate to 42", 42.0, formulaCell.getNumericCellValue(), 0.0001);
        }
    }

    /** T3 - ExtractorFactory service-provider path (ExtractorProvider -> POIXMLExtractorFactory). */
    @Test
    public void extractors() throws Exception {
        String xlsxText = ExtractorFactory.createExtractor(buildXlsx()).getText();
        Log.i(TAG, "xlsx extractor text: " + xlsxText);
        assertTrue("xlsx extractor text must contain 'Hello'", xlsxText.contains("Hello"));

        String xlsText = ExtractorFactory.createExtractor(buildHssf()).getText();
        Log.i(TAG, "xls extractor text: " + xlsText);
        assertTrue("xls extractor text must contain 'Hello'", xlsText.contains("Hello"));
    }

    /** T4 - direct proof the kept aalto StAX implementation is resolved under R8. */
    @Test
    public void staxIsAalto() throws Exception {
        XMLInputFactory factory = XMLInputFactory.newInstance();
        String impl = factory.getClass().getName();
        Log.i(TAG, "XMLInputFactory impl: " + impl);
        assertTrue("XMLInputFactory impl must be aalto but was: " + impl,
                impl.contains("aalto"));

        // diagnostic probe: which TransformerFactory resolves under minification?
        javax.xml.transform.TransformerFactory tf = javax.xml.transform.TransformerFactory.newInstance();
        Log.i(TAG, "TransformerFactory impl: " + tf.getClass().getName());
    }

    /**
     * T5 - pins the harness premise itself: the app's proguard-rules.pro blanket keep
     * ({@code -keep class m.co.rh.id.apoi_spreadsheet.** { *; }}) must keep library class
     * NAMES unchanged, because this test APK references library classes by name; the names
     * in T1-T4 would silently stop pointing at the real classes if the blanket keep was
     * weakened (allowshrinking/allowobfuscation added or the rule dropped). The aalto side
     * of the premise is already covered by T4.
     */
    @Test
    public void libraryClassesAreIdentityNamed() throws Exception {
        assertEquals(
                "library classes must stay identity-named (app proguard-rules.pro blanket keep)",
                "m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.WorkbookFactory",
                WorkbookFactory.class.getName());

        // NOTE: do NOT probe consumer-rule-pinned classes reflectively here (e.g.
        // Class.forName("org.apache.xmlbeans.impl.schema.TypeSystemHolder")). Standalone
        // init of that holder fails with SchemaTypeLoaderException ("Could not locate
        // compiled schema resource org/apache/xmlbeans/impl/schema/index.xsb") because
        // xmlbeans 5.3.0 ships no impl/schema/*.xsb resources - no real POI code path ever
        // triggers it (T1 passes). Its identity-name is guarded statically by
        // :r8-smoke:verifyR8Mapping instead.
    }
}
