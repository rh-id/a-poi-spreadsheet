// Derived from Apache POI (https://github.com/apache/poi @ commit 094968cfc3d48224db08f0b7f0a6fc341b035114); this file has been modified for Android compatibility by the a-poi-spreadsheet project.

package m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.crypt.temp;

import android.util.Log;

import org.apache.commons.io.output.CountingOutputStream;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import javax.crypto.Cipher;
import javax.crypto.CipherInputStream;
import javax.crypto.CipherOutputStream;
import javax.crypto.spec.SecretKeySpec;

import m.co.rh.id.apoi_spreadsheet.base.util.TempFile;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.crypt.ChainingMode;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.crypt.CipherAlgorithm;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.poifs.crypt.CryptoFunctions;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.Beta;
import m.co.rh.id.apoi_spreadsheet.org.apache.poi.util.RandomSingleton;


/**
 * EncryptedTempData can be used to buffer binary data in a secure way, by using encrypted temp files.
 */
@Beta
public class EncryptedTempData {

    private static final String TAG = "EncryptedTempData";
    private static final CipherAlgorithm cipherAlgorithm = CipherAlgorithm.aes128;
    private static final String PADDING = "PKCS5Padding";
    private final SecretKeySpec skeySpec;
    private final byte[] ivBytes;
    private final File tempFile;
    private CountingOutputStream outputStream;

    public EncryptedTempData() throws IOException {
        ivBytes = new byte[16];
        byte[] keyBytes = new byte[16];
        RandomSingleton.getInstance().nextBytes(ivBytes);
        RandomSingleton.getInstance().nextBytes(keyBytes);
        skeySpec = new SecretKeySpec(keyBytes, cipherAlgorithm.jceId);
        tempFile = TempFile.createTempFile("poi-temp-data", ".tmp");
    }

    /**
     * Returns the output stream for writing the data.<p>
     * Make sure to close it, otherwise the last cipher block is not written completely.
     *
     * @return the outputstream
     * @throws IOException if the writing to the underlying file fails
     */
    public OutputStream getOutputStream() throws IOException {
        Cipher ciEnc = CryptoFunctions.getCipher(skeySpec, cipherAlgorithm, ChainingMode.cbc, ivBytes, Cipher.ENCRYPT_MODE, PADDING);
        outputStream = new CountingOutputStream(new CipherOutputStream(Files.newOutputStream(tempFile.toPath()), ciEnc));
        return outputStream;
    }

    /**
     * Returns the input stream for reading the previously written encrypted data
     *
     * @return the inputstream
     * @throws IOException if the reading of the underlying file fails
     */
    public InputStream getInputStream() throws IOException {
        Cipher ciDec = CryptoFunctions.getCipher(skeySpec, cipherAlgorithm, ChainingMode.cbc, ivBytes, Cipher.DECRYPT_MODE, PADDING);
        return new CipherInputStream(Files.newInputStream(tempFile.toPath()), ciDec);
    }

    /**
     * @return number of bytes stored in the temp data file (the number you should expect after you decrypt the data)
     */
    public long getByteCount() {
        return outputStream == null ? 0 : outputStream.getByteCount();
    }

    /**
     * Removes the temporarily backing file
     */
    public void dispose() {
        if (!tempFile.delete()) {
            Log.w(TAG, tempFile.getAbsolutePath() + " can't be removed (or was already removed).");
        }
    }
}
