# Rules for the androidTest (release) APK only.
# Keep every test-APK class (AGP injects equivalent rules for minified test APKs; restated here for robustness).
-keep,allowobfuscation class ** { *; }
# Keep the fork's package names in the test APK so ServiceLoader provider wiring stays identity-named.
-keeppackagenames m.co.rh.id.apoi_spreadsheet.**
# androidx.test:monitor's Kotlin lambdas crashed with NoClassDefFoundError kotlin.jvm.internal.Lambda after R8 horizontal class merging; keep the stdlib.
-keep class kotlin.** { *; }
# androidx.test:monitor Tracer$Span references compile-only errorprone annotations.
-dontwarn com.google.errorprone.annotations.MustBeClosed
