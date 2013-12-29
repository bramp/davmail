package davmail.io;

import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Partial output stream, start at startIndex and write maxSize bytes.
 */
public final class PartialOutputStream extends FilterOutputStream {
    private int size;
    private int bufferSize;
    private final int startIndex;
    private final int maxSize;

    public PartialOutputStream(OutputStream os, int startIndex, int maxSize) {
        super(os);
        this.startIndex = startIndex;
        this.maxSize = maxSize;
    }

    @Override
    public void write(int b) throws IOException {
        size++;
        if ((size > startIndex) && (bufferSize < maxSize)) {
            super.write(b);
            bufferSize++;
        }
    }
}
