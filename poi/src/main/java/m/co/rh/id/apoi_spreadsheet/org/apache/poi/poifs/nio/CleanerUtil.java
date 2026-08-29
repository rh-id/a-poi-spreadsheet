/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project. The invocation uses plain reflection instead of MethodHandles because signature-polymorphic invoke blocks D8 dexing for apps with minSdk < 26.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.nio;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.ByteBuffer;
import java.security.AccessController;
import java.security.PrivilegedAction;

import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.ExceptionUtil;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.SuppressForbidden;


/**
 * This is taken from Hadoop at https://issues.apache.org/jira/browse/HADOOP-12760 and
 * https://github.com/apache/hadoop/blob/trunk/hadoop-common-project/hadoop-common/src/main/java/org/apache/hadoop/util/CleanerUtil.java
 * Unfortunately this is not available in some general utility library yet, but hopefully will be at some point.
 *
 * sun.misc.Cleaner has moved in OpenJDK 9 and
 * sun.misc.Unsafe#invokeCleaner(ByteBuffer) is the replacement.
 * This class is a hack to use sun.misc.Cleaner in Java 8 and
 * use the replacement in Java 9+.
 * This implementation is inspired by LUCENE-6989.
 */
@SuppressForbidden("uses java.security features deprecated in java 17 - no other option though")
public final class CleanerUtil {

    // Prevent instantiation
    private CleanerUtil(){}

    /**
     * <code>true</code>, if this platform supports unmapping mmapped files.
     */
    public static final boolean UNMAP_SUPPORTED;

    /**
     * if {@link #UNMAP_SUPPORTED} is {@code false}, this contains the reason
     * why unmapping is not supported.
     */
    public static final String UNMAP_NOT_SUPPORTED_REASON;


    private static final BufferCleaner CLEANER;

    /**
     * Reference to a BufferCleaner that does unmapping.
     * @return {@code null} if not supported.
     */
    public static BufferCleaner getCleaner() {
        return CLEANER;
    }

    static {
        final Object hack = AccessController.doPrivileged(
                (PrivilegedAction<Object>) CleanerUtil::unmapHackImpl);
        if (hack instanceof BufferCleaner) {
            CLEANER = (BufferCleaner) hack;
            UNMAP_SUPPORTED = true;
            UNMAP_NOT_SUPPORTED_REASON = null;
        } else {
            CLEANER = null;
            UNMAP_SUPPORTED = false;
            UNMAP_NOT_SUPPORTED_REASON = hack.toString();
        }
    }

    @SuppressForbidden("Java 9 Jigsaw allows access to sun.misc.Cleaner, so setAccessible works")
    private static Object unmapHackImpl() {
        try {
            try {
                // *** sun.misc.Unsafe unmapping (Java 9+) ***
                final Class<?> unsafeClass = Class.forName("sun.misc.Unsafe");
                // first check if Unsafe has the right method, otherwise we can
                // give up without doing any security critical stuff:
                final Method unmapper = unsafeClass.getMethod(
                        "invokeCleaner", ByteBuffer.class);
                // fetch the unsafe instance and pass it as receiver to the method:
                final Field f = unsafeClass.getDeclaredField("theUnsafe");
                f.setAccessible(true);
                final Object theUnsafe = f.get(null);
                return newBufferCleaner(ByteBuffer.class, theUnsafe, unmapper, null);
            } catch (SecurityException se) {
                // rethrow to report errors correctly (we need to catch it here,
                // as we also catch RuntimeException below!):
                throw se;
            } catch (ReflectiveOperationException | RuntimeException e) {
                // *** sun.misc.Cleaner unmapping (Java 8) ***
                final Class<?> directBufferClass =
                        Class.forName("java.nio.DirectByteBuffer");

                final Method cleanerMethod = directBufferClass.getMethod("cleaner");
                cleanerMethod.setAccessible(true);
                final Class<?> cleanerClass = cleanerMethod.getReturnType();

                /*
                 * The invocation below is basically equivalent to the
                 * following code:
                 *
                 * void unmapper(ByteBuffer byteBuffer) {
                 *   sun.misc.Cleaner cleaner =
                 *       ((java.nio.DirectByteBuffer) byteBuffer).cleaner();
                 *   if (Objects.nonNull(cleaner)) {
                 *     cleaner.clean();
                 *   }
                 * }
                 */
                final Method cleanMethod = cleanerClass.getMethod("clean");
                return newBufferCleaner(directBufferClass, null, cleanerMethod, cleanMethod);
            }
        } catch (SecurityException se) {
            return "Unmapping is not supported, because not all required " +
                    "permissions are given to the Hadoop JAR file: " + se +
                    " [Please grant at least the following permissions: " +
                    "RuntimePermission(\"accessClassInPackage.sun.misc\") " +
                    " and ReflectPermission(\"suppressAccessChecks\")]";
        } catch (ReflectiveOperationException | RuntimeException e) {
            return "Unmapping is not supported on this platform, " +
                    "because internal Java APIs are not compatible with " +
                    "this Hadoop version: " + e;
        }
    }

    private static BufferCleaner newBufferCleaner(
            final Class<?> unmappableBufferClass, final Object receiver,
            final Method unmapperMethod, final Method cleanMethod) {
        return buffer -> {
            if (!buffer.isDirect()) {
                throw new IllegalArgumentException(
                        "unmapping only works with direct buffers");
            }
            if (!unmappableBufferClass.isInstance(buffer)) {
                throw new IllegalArgumentException("buffer is not an instance of " +
                        unmappableBufferClass.getName());
            }
            final Throwable error = AccessController.doPrivileged(
                    (PrivilegedAction<Throwable>) () -> {
                        try {
                            if (receiver != null) {
                                unmapperMethod.invoke(receiver, buffer);
                            } else {
                                final Object cleaner = unmapperMethod.invoke(buffer);
                                if (cleaner != null) {
                                    cleanMethod.invoke(cleaner);
                                }
                            }
                        } catch (Throwable t) {
                            if (ExceptionUtil.isFatal(t)) {
                                ExceptionUtil.rethrow(t);
                            }
                        }
                        return null;
                    });
            if (error != null) {
                if (ExceptionUtil.isFatal(error)) {
                    ExceptionUtil.rethrow(error);
                }
                throw new IOException("Unable to unmap the mapped buffer", error);
            }
        };
    }

    /**
     * Pass in an implementation of this interface to cleanup ByteBuffers.
     * CleanerUtil implements this to allow unmapping of bytebuffers
     * with private Java APIs.
     */
    @FunctionalInterface
    public interface BufferCleaner {
        void freeBuffer(ByteBuffer b) throws IOException;
    }
}
