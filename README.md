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
1. XSSF Module (`.xlsx` spreadsheet file) and its dependency are copied and adapted into this project (source trees under `poi/` and `poi-ooxml/`, published as one artifact since v0.1.0)
2. Apache Log4J will be replaced by android Log classes
3. Add extra Android-compatibility classes (Android context bridge, `java.awt` replacements) — since v0.1.0 these ship inside the single `a-poi-spreadsheet` module (source tree under `base/`)
4. Quiet a number of files modified and adapted for Android including Test files for Android Instrumented Test
5. Some portion of the functionalities are not really migrated yet (see `FIXME` comments)

## Using this library

`minSdk` is 26 (due to required Java 8 Compatibility). The library is declared and tested on API 26+.

- **App `minSdk` 26 or higher**: nothing extra is needed. The manifest merges normally, no runtime guard is required (`java.time` / `java.nio.file` are always present), and the bundled `consumer-rules.pro` are picked up automatically when your app enables minification (see [Proguard Configuration](#proguard-configuration)).
- **App `minSdk` below 26**: additional setup is required, see the next subsection.

### Using this library in apps with `minSdk` below 26

Apps with a lower `minSdk` (e.g. 21) can still include this library. The library dexes fine below API 26 (since v0.0.7, verified with D8/R8 full mode at `minSdk 21`), but its code requires APIs that only exist from API 26 onwards (`java.time`, `java.nio.file`). On older devices the app must simply not execute any POI code — Android loads classes lazily, so library classes sitting unused in the APK are harmless:

1. Add `xmlns:tools="http://schemas.android.com/tools"` and `<uses-sdk tools:overrideLibrary="m.co.rh.id.apoi_spreadsheet" />` to your manifest (required because the library declares `minSdk 26`).
2. Isolate all POI usage in a wrapper class and only instantiate it inside a `Build.VERSION.SDK_INT >= 26` guard, e.g.:
```
if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
    spreadsheetHelper.open(file); // SpreadsheetHelper is the only class touching POI
} else {
    // spreadsheet features unavailable on this device
}
```
Never trigger the guarded path on older devices — a single unguarded POI call will crash with `NoClassDefFoundError` (e.g. `java.nio.file.OpenOption`).

Actually running POI on API 21-25 is not supported or tested. If you attempt it anyway, plain core library desugaring is not enough (commons-io requires `java.nio.file`); you would need `coreLibraryDesugaring "com.android.tools:desugar_jdk_libs_nio:2.1.5"` — treat this as best-effort at your own risk.

This project supports JitPack. To use it, add the jitpack repository to your `settings.gradle`:
```
// settings.gradle
dependencyResolutionManagement {
    repositories {
        google()
        mavenCentral()
        maven { url "https://jitpack.io" }
    }
}
```

Older setups that still declare repositories in the root `build.gradle` via `allprojects` may add jitpack there instead.

Include this to your module dependency (module build.gradle)
```
dependencies {
    implementation 'com.github.rh-id:a-poi-spreadsheet:v0.1.0'
}
```

**Migrating from v0.0.7 or earlier:** those versions published 3 artifacts (`a-poi-spreadsheet`, `a-poi-spreadsheet-base`, `a-poi-spreadsheet-ooxml`). From v0.1.0 everything is consolidated into the single `a-poi-spreadsheet` artifact — delete the `-base` and `-ooxml` dependency lines when upgrading. The old coordinates remain available on JitPack at v0.0.7 only.

Note: since v0.1.0 is a single artifact, apps that only process legacy `.xls` (HSSF) workbooks will also pull in the OOXML dependencies (`poi-ooxml-full` schemas, XMLBeans, xmlsec, etc.). If APK size matters for such apps, staying on v0.0.7 with the 3-artifact setup is an option.

Set application context during `Application.onCreate` or before using it to poi spreadsheet context: `POISpreadsheetContext.getInstance().setAppContext(Context)`

`POISpreadsheetContext` is not from original Apache POI, it was used to bridge Android context and Apache POI execution.

`POISpreadsheetContext` implements `ExecutorService` in hope that you will use this context to execute any of the Apache POI operation.

## Proguard Configuration

No manual rules are required. The published AAR bundles its `consumer-rules.pro` (via `consumerProguardFiles`), so R8/ProGuard picks them up automatically when your app enables minification. This repo itself verifies the shipped rules via the `:r8-smoke` minified app harness and its `verifyR8Mapping` task, which asserts that the consumer-rule-pinned classes stay identity-named under R8 full mode.

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

- **Required:** `xmlsec` and `log4j-api` both ship a Maven `META-INF/DEPENDENCIES` file, which makes consumer builds fail at `mergeReleaseJavaResource` ("2 files found with path 'META-INF/DEPENDENCIES'"). Exclude it in your app's `android` block:

```
android {
    packaging {
        resources.excludes += "META-INF/DEPENDENCIES"
    }
}
```

- The schema classes ship in `poi-ooxml-full`, which contains ~9000 `.xsb` binary resources — they must reach your APK. Default AGP packaging keeps them, so nothing to do unless you have custom packaging filters that strip them.

## Licenses

 This project is an independent Android adaptation of [Apache POI](https://poi.apache.org/) (fork point: [apache/poi@094968cf](https://github.com/apache/poi/commit/094968cfc3d48224db08f0b7f0a6fc341b035114), tag REL_5_5_1). It is **not** an official Apache POI product.

- Upstream Apache POI code remains Copyright 2003-2025 The Apache Software Foundation, licensed under the [Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0). The upstream LICENSE and NOTICE texts are preserved in [`poi_legal`](poi_legal) and shipped inside the published artifact under `META-INF/LICENSE-a-poi-spreadsheet.txt` and `META-INF/NOTICE-a-poi-spreadsheet.txt`.
- Modifications and additions for Android are Copyright 2024-2026 Ruby Hartono, licensed under the Apache License, Version 2.0 (see [LICENSE](LICENSE)).
- The `javax.xml.crypto:jsr105-api` dependency used for XML digital signatures is Sun Microsystems code dual-licensed under CDDL / GPLv2 with Classpath Exception. Linking against it does not affect your licensing.

**Downstream projects may license their own applications/libraries under any terms.** Using this library only incurs the attribution obligations listed above and in the packaged NOTICE.

See the [GitHub Releases](https://github.com/rh-id/a-poi-spreadsheet/releases) page for the changelog of each version.

