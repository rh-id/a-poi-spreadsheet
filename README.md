# a-poi-spreadsheet

![Languages](https://img.shields.io/github/languages/top/rh-id/a-poi-spreadsheet)
![JitPack](https://img.shields.io/jitpack/v/github/rh-id/a-poi-spreadsheet)
![Downloads](https://jitpack.io/v/rh-id/a-poi-spreadsheet/week.svg)
![Downloads](https://jitpack.io/v/rh-id/a-poi-spreadsheet/month.svg)
![Android CI](https://github.com/rh-id/a-poi-spreadsheet/actions/workflows/gradlew-build.yml/badge.svg)
![Emulator Test](https://github.com/rh-id/a-poi-spreadsheet/actions/workflows/android-emulator-test.yml/badge.svg)

This is android library project that copied,import, and adapted from Apache POI.
The latest commit since: https://github.com/apache/poi/commit/6a8994ee0e6c59aa231570307a5dd213784993c3

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
    implementation 'com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet:v0.0.4'
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-base:v0.0.4"
    implementation "com.github.rh-id.a-poi-spreadsheet:a-poi-spreadsheet-ooxml:v0.0.4"
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

This project is an independent Android adaptation of [Apache POI](https://poi.apache.org/) (fork point: [apache/poi@6a8994e](https://github.com/apache/poi/commit/6a8994ee0e6c59aa231570307a5dd213784993c3)). It is **not** an official Apache POI product.

- Upstream Apache POI code remains Copyright 2003-2024 The Apache Software Foundation, licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0). The upstream LICENSE and NOTICE texts are preserved in [`poi_legal`](poi_legal) and shipped inside each published artifact under `META-INF/LICENSE-a-poi-spreadsheet.txt` and `META-INF/NOTICE-a-poi-spreadsheet.txt`.
- Modifications and additions for Android are Copyright 2024 Ruby Hartono, licensed under the Apache License, Version 2.0 (see [LICENSE](LICENSE)).
- The `javax.xml.crypto:jsr105-api` dependency used for XML digital signatures is Sun Microsystems code dual-licensed under CDDL / GPLv2 with Classpath Exception. Linking against it does not affect your licensing.

**Downstream projects may license their own applications/libraries under any terms.** Using this library only incurs the attribution obligations listed above and in the packaged NOTICE.

### Changelog

#### v0.0.4
- Added Maven POM license/developer/scm metadata to all published artifacts
- LICENSE and NOTICE now packaged inside artifacts (`META-INF/`)
- Disclosed third-party licensing for `jsr105-api` (CDDL/GPLv2+Classpath Exception)
- Added modification notices and license headers per Apache-2.0 §4 requirements

