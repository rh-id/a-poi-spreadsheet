# a-poi-spreadsheet consumer rules.
# Applied automatically to consumer apps via consumerProguardFiles.
#
# Merged from the former per-module consumer rules when the base/poi/poi-ooxml
# modules were consolidated into this single module (v0.1.0):
# - the former base module had no rules (all its classes are statically reachable);
# - the former poi and poi-ooxml modules contributed the rules below, with the
#   ServiceLoader provider keeps unified into a single section.

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

# --- XMLBeans type systems ---
# poi-ooxml-full 5.5.1 ships its type system as
# org.apache.poi.schemas.ooxml.system.ooxml.TypeSystemHolder (the old
# schemaorg_apache_xmlbeans.system.* packages no longer exist).
-keep class org.apache.poi.schemas.** { *; }
-keep class org.apache.xmlbeans.** { *; }

# --- Schema classes ---
# XMLBeans instantiates schema impl classes reflectively:
# Class.forName(fullJavaImplName) + getDeclaredConstructor(SchemaType, boolean)
# (SchemaTypeImpl.java) and enum value classes via Class.forName(name + "$Enum").
-keep class org.openxmlformats.schemas.** { *; }
-keep class com.microsoft.schemas.** { *; }
-keep class org.etsi.** { *; }
# NOTE: org.w3c.dom must be kept by name: the DOM API copy shipped by xml-apis is
# invoked through the framework's org.w3c.dom.Node (bootclasspath) - if R8 renames
# it, xerces' DOM impl mixes renamed/identity signatures and every xlsx save fails
# (NoSuchMethodError getOwnerDocument). "org.w3.**" does NOT match "org.w3c.dom.**"
# (pattern is a literal prefix, so this needs the explicit "org.w3c." prefix).
-keep class org.w3c.** { *; }

# --- XML-DSig provider loaded by name (SignatureInfo.findProvider) ---
-keep class org.apache.jcp.xml.dsig.internal.dom.XMLDSigRI { *; }
-dontwarn org.jcp.xml.dsig.internal.dom.**

# --- ServiceLoader providers (pinned explicitly; R8 models ServiceLoader.load
#     with constant literals, but this stays robust across R8 versions/modes).
#     Covers the HSSF (former poi module) and XSSF/OOXML (former poi-ooxml module)
#     provider factories; their META-INF/services registrations are merged in this
#     single artifact. ---
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.WorkbookProvider
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.ExtractorProvider
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.hssf.usermodel.HSSFWorkbookFactory { <init>(); }
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbookFactory { <init>(); }
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.MainExtractorFactory { <init>(); }
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.ooxml.extractor.POIXMLExtractorFactory { <init>(); }

# --- Optional integrations of dependencies (absent on Android, never used at runtime) ---
# log4j-api (API-only, no backend) references optional aQute.bnd annotations
-dontwarn aQute.bnd.annotation.**

# commons-compress optional codec providers (not shipped, absent on Android)
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

# xmlbeans schema-compiler code-gen tooling (org.apache.xmlbeans.impl.config.** -> javaparser)
-dontwarn com.github.javaparser.**
# xmlbeans optional Saxon XPath/XQuery engine (org.apache.xmlbeans.impl.xpath.saxon.**)
-dontwarn net.sf.saxon.**
# xmlbeans Maven-plugin integration (MavenPluginResolver, XMLBeansMojo)
-dontwarn org.apache.maven.**
# xmlbeans Ant task integration (org.apache.xmlbeans.impl.tool.XMLBean)
-dontwarn org.apache.tools.ant.**
# xmlbeans MavenPluginResolver catalog-resolver fallback (JDK-internal org.apache.xml.resolver)
-dontwarn com.sun.org.apache.xml.internal.resolver.**
# log4j-api (API-only, no backend) compile-time-only references:
# - edu.umd.cs.findbugs.annotations (nullability)
# - com.google.errorprone.annotations (InlineMe migration hints)
# - org.osgi.framework (optional OSGi runtime probe, OsgiServiceLocator)
# - org.osgi.annotation.* (OSGi bundle/versioning annotations on package-info/Activator)
-dontwarn edu.umd.cs.findbugs.annotations.**
-dontwarn com.google.errorprone.annotations.**
-dontwarn org.osgi.framework.**
-dontwarn org.osgi.annotation.**
# log4j-api nullability annotations (org.apache.logging.log4j.simple.internal.SimpleProvider)
-dontwarn org.jspecify.annotations.**
# slf4j-api 1.x's org.slf4j.LoggerFactory/IMarkerFactory/MDCAdapter reference their
# optional 1.x binding classes in org.slf4j.impl (StaticLoggerBinder,
# StaticMarkerBinder, StaticMDCBinder)
-dontwarn org.slf4j.impl.**

# xmlsec references javax.naming (absent on Android)
-dontwarn javax.naming.**
# xmlsec pulls in woodstox-core: its shaded MSV DTD-validation support references a
# command-line driver entry point (com.ctc.wstx.shaded.msv_core.driver.textui.Driver)
# that is not shipped; DTD validation is never used on Android
-dontwarn com.ctc.wstx.**
# xmlsec's jakarta.xml.bind support references jakarta.activation.DataHandler (MTOM
# attachments); JAXB/activation are absent on Android and never used by POI's DSig path
-dontwarn jakarta.activation.**
