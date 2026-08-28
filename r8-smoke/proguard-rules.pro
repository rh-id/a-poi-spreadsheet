# Rules for the app (release) APK only. The library consumer-rules.pro under test are merged in
# automatically from poi/consumer-rules.pro and poi-ooxml/consumer-rules.pro via consumerProguardFiles.
# androidx.test:monitor (test APK, deduped against this app APK) needs the Kotlin stdlib at runtime under R8 full mode.
-keep class kotlin.** { *; }
# Harness app has no POI entry points of its own, so without this R8 shrinks the whole library away; consumer rules still govern every dynamic path (aalto StAX, xmlbeans schemas, ServiceLoader factories, XML-DSig).
-keep class m.co.rh.id.apoi_spreadsheet.** { *; }
