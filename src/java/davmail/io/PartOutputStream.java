package davmail.io;

import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Filter to output only headers, also count full size
 */
public final class PartOutputStream extends FilterOutputStream {
    private static final int START = 0;
    private static final int CR = 1;
    private static final int CRLF = 2;
    private static final int CRLFCR = 3;
    private static final int BODY = 4;

    private int state = START;
    private int size;
    private int bufferSize;
    private final boolean writeHeaders;
    private final boolean writeBody;
    private final int startIndex;
    private final int maxSize;

    public PartOutputStream(OutputStream os, boolean writeHeaders, boolean writeBody,
                     int startIndex, int maxSize) {
        super(os);
        this.writeHeaders = writeHeaders;
        this.writeBody = writeBody;
        this.startIndex = startIndex;
        this.maxSize = maxSize;
    }

    @Override
    public void write(int b) throws IOException {
        size++;
        if (((state != BODY && writeHeaders) || (state == BODY && writeBody)) &&
                (size > startIndex) && (bufferSize < maxSize)
                ) {
            super.write(b);
            bufferSize++;
        }
        if (state == START) {
            if (b == '\r') {
                state = CR;
            }
        } else if (state == CR) {
            if (b == '\n') {
                state = CRLF;
            } else {
                state = START;
            }
        } else if (state == CRLF) {
            if (b == '\r') {
                state = CRLFCR;
            } else {
                state = START;
            }
        } else if (state == CRLFCR) {
            if (b == '\n') {
                state = BODY;
            } else {
                state = START;
            }
        }
    }
}
