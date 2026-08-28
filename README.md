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
    implementation 'com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet:v0.0.6'
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-base:v0.0.6"
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-ooxml:v0.0.6"
}
```

Set application context during `Application.onCreate` or before using it to poi spreadsheet context: `POISpreadsheetContext.getInstance().setAppContext(Context)`

`POISpreadsheetContext` is not from original Apache POI, it was used to bridge Android context and Apache POI execution.

`POISpreadsheetContext` implements `ExecutorService` in hope that you will use this context to execute any of the Apache POI operation.

## Proguard Configuration

No manual rules are required. Every published AAR bundles its own `consumer-rules.pro` (via `consumerProguardFiles`), so R8/ProGuard picks them up automatically when your app enables minification.

The bundled rules cover everything reflection makes invisible to R8:

- The StAX implementation (`aalto-xml` + `stax2-api`, plus Xerces' `XMLEventFactoryImpl`), which the JDK service lookup resolves by parsing `META-INF/services` text and calling `Class.forName` — without it every `.xlsx` read/write fails under minification.
- The XMLBeans type system (`org.apache.poi.schemas.**`, `org.apache.xmlbeans.**`) and the OOXML schema packages (`org.openxmlformats.**`, `com.microsoft.schemas.**`, `org.etsi.**`, `org.w3c.**`), whose `*Impl` and `$Enum` classes are instantiated reflectively.
- The reflection-loaded XML-DSig security provider (`XMLDSigRI`) and the optional Xerces `SecurityManager` hardening hook.
- The `ServiceLoader` provider factories (`HSSFWorkbookFactory`, `XSSFWorkbookFactory`, `MainExtractorFactory`, `POIXMLExtractorFactory`).

Optional manual additions for consumer apps:

```
# Only needed if your app adds the xml-resolver artifact itself; xmlbeans' own
# catalog-resolver fallback (JDK-internal com.sun.org.apache.xml.internal.resolver)
# is already covered by the bundled rules
-dontwarn org.apache.xml.resolver.**
```

If your app adds BouncyCastle itself (for encrypted workbook support), also keep its provider, since POI loads it by name (`CryptoFunctions`):

```
-keep class org.bouncycastle.jce.provider.BouncyCastleProvider { *; }
```

Packaging notes:

- The schema classes ship in `poi-ooxml-full`, which contains ~9000 `.xsb` binary resources — they must reach your APK. Default AGP packaging keeps them, so nothing to do unless you have custom packaging filters that strip them.
- Both `aalto-xml` and `xercesImpl` ship a `META-INF/services/javax.xml.stream.XMLEventFactory` entry. If the duplicate-resource build error appears, resolve it with either file — both factory implementations are kept, so either resolution works:

```
android {
    packaging {
        resources.pickFirsts += "META-INF/services/javax.xml.stream.XMLEventFactory"
    }
}
```

## Licenses

 This project is an independent Android adaptation of [Apache POI](https://poi.apache.org/) (fork point: [apache/poi@094968cf](https://github.com/apache/poi/commit/094968cfc3d48224db08f0b7f0a6fc341b035114), tag REL_5_5_1). It is **not** an official Apache POI product.

- Upstream Apache POI code remains Copyright 2003-2025 The Apache Software Foundation, licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0). The upstream LICENSE and NOTICE texts are preserved in [`poi_legal`](poi_legal) and shipped inside each published artifact under `META-INF/LICENSE-a-poi-spreadsheet.txt` and `META-INF/NOTICE-a-poi-spreadsheet.txt`.
- Modifications and additions for Android are Copyright 2024-2026 Ruby Hartono, licensed under the Apache License, Version 2.0 (see [LICENSE](LICENSE)).
- The `javax.xml.crypto:jsr105-api` dependency used for XML digital signatures is Sun Microsystems code dual-licensed under CDDL / GPLv2 with Classpath Exception. Linking against it does not affect your licensing.

**Downstream projects may license their own applications/libraries under any terms.** Using this library only incurs the attribution obligations listed above and in the packaged NOTICE.

See the [GitHub Releases](https://github.com/rh-id/a-poi-spreadsheet/releases) page for the changelog of each version.

