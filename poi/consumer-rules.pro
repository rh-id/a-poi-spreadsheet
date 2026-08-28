# a-poi-spreadsheet (poi module) consumer rules.
# Applied automatically to consumer apps via consumerProguardFiles.

# --- StAX implementation (aalto-xml + stax2-api) ---
# javax.xml.stream does not exist on Android; xml-apis' FactoryFinder resolves the
# implementation by parsing META-INF/services text and Class.forName(name), which is
# invisible to R8. Keep the implementations or every .xlsx read/write fails.
-keep class com.fasterxml.aalto.stax.** { *; }
-keep class org.codehaus.stax2.** { *; }
# xercesImpl also ships a javax.xml.stream.XMLEventFactory service entry
-keep class org.apache.xerces.stax.XMLEventFactoryImpl { *; }

# xercesImpl JAXP providers (META-INF/services/javax.xml.parsers.* and
# javax.xml.datatype.* entries) are resolved by name through FactoryFinder inside
# POI's XMLHelper/DocumentHelper ("Provider org.apache.xerces.jaxp.DocumentBuilderFactoryImpl
# not found" under R8)
-keep class org.apache.xerces.jaxp.** { *; }
# xerces parser core is loaded by name via its internal ObjectFactory
# ("Provider org.apache.xerces.parsers.XIncludeAwareParserConfiguration not found" under R8)
-keep class org.apache.xerces.parsers.** { *; }
# xerces datatype-validator factories are loaded by name too
# (DTDDVFactory.getInstance -> impl.dv.dtd.DTDDVFactoryImpl,
#  SchemaDVFactory.getInstance -> impl.dv.xs.SchemaDVFactoryImpl)
-keep class org.apache.xerces.impl.dv.** { *; }

# --- xml-apis API trees must keep their identity names ---
# xercesImpl drags in xml-apis, which ships copies of the javax.xml.*,
# org.xml.sax and org.w3c.dom API classes that also exist in the Android
# framework. If R8 renames the copies, POI's references are rewired to the
# renamed API and framework implementations loaded by service lookups fail
# casts/instanceof (ClassCastException in XMLHelper.getTransformerFactory).
# keepnames makes the framework versions win resolution exactly as in
# non-minified builds (unused duplicates may still be removed).
-keep class javax.xml.** { *; }
-keep class org.xml.sax.** { *; }
-keepnames class org.w3c.dom.** { *; }

# NOTE on Transformers: Android supplies Xalan (org.apache.xalan.*) inside its
# framework image, and javax.xml.** API classes are kept identity-named above, so
# TransformerFactory.newInstance() resolves the framework implementation exactly as
# in non-minified builds. No rule is needed for the transform stack itself.

# --- Optional Xerces entity-expansion hardening hook (loaded reflectively by XMLHelper) ---
-keep class org.apache.xerces.util.SecurityManager { *; }

# --- HSSF ServiceLoader providers (pinned explicitly; R8 models ServiceLoader.load
#     with constant literals, but this stays robust across R8 versions/modes) ---
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.WorkbookProvider
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.ExtractorProvider
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbookFactory { <init>(); }
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.MainExtractorFactory { <init>(); }

# log4j-api (API-only, no backend) references optional aQute.bnd annotations
-dontwarn aQute.bnd.annotation.**

# --- commons-compress optional codec providers (not shipped, absent on Android) ---
# org.apache.commons.compress.compressors.CompressorStreamFactory's codec registry
# references zstd-jni (com.github.luben.zstd), xz (org.tukaani.xz) and brotli
# (org.brotli.dec) providers
-dontwarn com.github.luben.zstd.**
-dontwarn org.tukaani.xz.**
-dontwarn org.brotli.dec.**
# commons-compress's harmony pack200 helpers reference ObjectWeb ASM (absent on
# Android; pack200 is JDK-tooling only) [found via R8-full instrumented suite run]
-dontwarn org.objectweb.asm.**
# commons-math3 geometry (Line.getTransform) references java.awt.geom, absent on
# Android; the spreadsheet formula path never touches geometry partitioning
-dontwarn java.awt.geom.**
# commons-lang3 MethodInvokers.asInterfaceInstance references
# java.lang.invoke.MethodHandleProxies, absent from Android's java.lang.invoke
-dontwarn java.lang.invoke.MethodHandleProxies

# xerces' XMLCatalogResolver references the optional xml-commons-resolver library
# (org.apache.xml.resolver.*) which is not shipped; catalog resolution is an
# optional DTD/entity-catalog feature never used on Android
-dontwarn org.apache.xml.resolver.**
