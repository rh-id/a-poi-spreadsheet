# a-poi-spreadsheet (poi-ooxml module) consumer rules.
# Applied automatically to consumer apps via consumerProguardFiles.

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

# --- XSSF ServiceLoader providers (pinned explicitly; see poi module rules) ---
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.ss.usermodel.WorkbookProvider
-keepnames class m.co.rh.id.apoi_spreadsheet.org.apache.poi.extractor.ExtractorProvider
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.xssf.usermodel.XSSFWorkbookFactory { <init>(); }
-keep class m.co.rh.id.apoi_spreadsheet.org.apache.poi.ooxml.extractor.POIXMLExtractorFactory { <init>(); }

# --- Optional integrations of this module's dependencies (absent on Android, never used at runtime) ---
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
