# a-poi-spreadsheet

![Languages](https://img.shields.io/github/languages/top/rh-id/a-poi-spreadsheet)
![JitPack](https://img.shields.io/jitpack/v/github/rh-id/a-poi-spreadsheet)
![Downloads](https://jitpack.io/v/rh-id/a-poi-spreadsheet/week.svg)
![Downloads](https://jitpack.io/v/rh-id/a-poi-spreadsheet/month.svg)
![Android CI](https://github.com/rh-id/a-poi-spreadsheet/actions/workflows/gradlew-build.yml/badge.svg)
![Emulator Test](https://github.com/rh-id/a-poi-spreadsheet/actions/workflows/android-emulator-test.yml/badge.svg)

This is android library project that copied,import, and adapted from Apache POI.
The latest commit since: https://github.com/apache/poi/commit/094968cfc3d48224db08f0b7f0a6fc341b035114 (tag `REL_5_5_1`, POI 5.5.1)

From that point, XSSF module was cut off and adapted to this library

POI license notice is in `poi_legal`

NOTE: This is not official library or project from Apache POI team.

## Changes Adapted
1. XSSF Module (`.xlsx` spreadsheet file) and its dependency are copied and moved to poi modules here with new workspaces
2. Apache Log4J will be replaced by android Log classes
3. Add extra classes at base module for compatibility
4. Quiet a number of files modified and adapted for Android including Test files for Android Instrumented Test
5. Some portion of the functionalities are not really migrated yet (see `FIXME` comments)

## Using this library

`minSdk` is 26 (due to required Java 8 Compatibility).

This project support jitpack, in order to use this, you need to add jitpack to your project root build.gradle:
```
allprojects {
    repositories {
        google()
        mavenCentral()
        maven { url "https://jitpack.io" }
    }
}
```

Include this to your module dependency (module build.gradle)
```
dependencies {
    implementation 'com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet:v0.0.5'
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-base:v0.0.5"
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-ooxml:v0.0.5"
}
```

Set application context during `Application.onCreate` or before using it to poi spreadsheet context: `POISpreadsheetContext.getInstance().setAppContext(Context)`

`POISpreadsheetContext` is not from original Apache POI, it was used to bridge Android context and Apache POI execution.

`POISpreadsheetContext` implements `ExecutorService` in hope that you will use this context to execute any of the Apache POI operation.

## Proguard Configuration

```
-dontwarn org.apache.**
-dontwarn org.openxmlformats.schemas.**
-dontwarn org.etsi.**
-dontwarn org.w3.**
-dontwarn com.microsoft.schemas.**
-dontwarn com.graphbuilder.**
-dontwarn aQute.bnd.annotation.spi.ServiceProvider
-dontnote org.apache.**
-dontnote org.openxmlformats.schemas.**
-dontnote org.etsi.**
-dontnote org.w3.**
-dontnote com.microsoft.schemas.**
-dontnote com.graphbuilder.**

-keeppackagenames org.apache.poi.ss.formula.function

-keep class org.apache.logging.** { *; }
-keep class org.apache.commons.** { *; }
-keep class org.apache.xmlbeans.** { *; }
-keep class org.openxmlformats.schemas.** { *; }
-keep class com.microsoft.schemas.** { *; }
-keep class javax.xml.** { *; }

-keep class schemaorg_apache_xmlbeans.system.sF1327CCA741569E70F9CA8C9AF9B44B2.TypeSystemHolder { public final static *** typeSystem; }
```

## Licenses

 This project is an independent Android adaptation of [Apache POI](https://poi.apache.org/) (fork point: [apache/poi@094968cf](https://github.com/apache/poi/commit/094968cfc3d48224db08f0b7f0a6fc341b035114), tag REL_5_5_1). It is **not** an official Apache POI product.

- Upstream Apache POI code remains Copyright 2003-2025 The Apache Software Foundation, licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0). The upstream LICENSE and NOTICE texts are preserved in [`poi_legal`](poi_legal) and shipped inside each published artifact under `META-INF/LICENSE-a-poi-spreadsheet.txt` and `META-INF/NOTICE-a-poi-spreadsheet.txt`.
- Modifications and additions for Android are Copyright 2024-2026 Ruby Hartono, licensed under the Apache License, Version 2.0 (see [LICENSE](LICENSE)).
- The `javax.xml.crypto:jsr105-api` dependency used for XML digital signatures is Sun Microsystems code dual-licensed under CDDL / GPLv2 with Classpath Exception. Linking against it does not affect your licensing.

**Downstream projects may license their own applications/libraries under any terms.** Using this library only incurs the attribution obligations listed above and in the packaged NOTICE.

### Changelog

#### v0.0.5
- Updated upstream base from apache/poi@6a8994e (5.2.5 era, Feb 2024) to [REL_5_5_1](https://github.com/apache/poi/commit/094968cfc3d48224db08f0b7f0a6fc341b035114) (POI 5.5.1, Nov 2025) - 238 carried files regenerated, picked up ~2 years of upstream fixes and improvements
- Instrumented tests refreshed from upstream: all 39 upstream-changed test files ported to JUnit 4 style (32 with material changes, 7 already in sync) + 6 new tests added (HSSFParser, XSSFParser, HSSFRowCopyRowFrom, WorkdayFunc, CellUtil, OutOfOrderColumns); new upstream classes HSSFParser/XSSFParser + HSSFReadException/XSSFReadException carried into main sources
- Dependencies updated to upstream 5.5.1 baseline: poi-ooxml-full 5.5.1, commons-io 2.21.0, commons-collections4 4.5.0, commons-compress 1.28.0, commons-codec 1.20.0, xmlsec 3.0.6
- Carried previously-stripped classes back from upstream where Android-safe: `CellPropertyType`, `CellPropertyCategory`, `poifs/nio/CleanerUtil`, `util/POIException`, `openxml4j/opc/OPCComplianceFlags`, `ReferenceRelationship`, `HyperlinkRelationship`, `Reproducibility`, `InvalidZipException` and more
- Restored fork's android.graphics.Color APIs on `XSSFColor`, `XSSFTextRun`, `XSSFTextParagraph`
- Breaking changes vs v0.0.4:
  - `XSSFTextRun#getFontColor`/`setFontColor`, `XSSFTextParagraph#getBulletFontColor`/`setBulletFontColor` use `android.graphics.Color` (unchanged from v0.0.4, but signatures differ from upstream POI)
  - `SignatureConfig#getTspService`/`setTspService` removed (upstream default `TSPTimeStampService` depends on BouncyCastle, not carried)
  - `XSSFColor(java.awt.Color, IndexedColorMap)`-style awt constructors remain unavailable (use `android.graphics.Color` or `byte[]` variants)
  - `ss.usermodel.CellStyle#getFormatProperties`/`invalidateCachedProperties` now available again (upstream API restored)
- xmlbeans schemas now come from poi-ooxml-full 5.5.1 - the proguard keep rule for `schemaorg_apache_xmlbeans.system.*` in consumer projects remains valid
- On-device fixes: formula evaluation now loads function metadata from Android assets; depends on log4j-api (API-only, no logging backend) for XMLBeans 5.x compatibility

#### v0.0.4
- Added Maven POM license/developer/scm metadata to all published artifacts
- LICENSE and NOTICE now packaged inside artifacts (`META-INF/`)
- Disclosed third-party licensing for `jsr105-api` (CDDL/GPLv2+Classpath Exception)
- Added modification notices and license headers per Apache-2.0 §4 requirements

