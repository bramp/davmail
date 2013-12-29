package davmail.io;

import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Filter to limit output lines to max body lines after header
 */
final public class TopOutputStream extends FilterOutputStream {
    private static final int START = 0;
    private static final int CR = 1;
    private static final int CRLF = 2;
    private static final int CRLFCR = 3;
    private static final int BODY = 4;

    private int maxLines;
    private int state = START;

    public TopOutputStream(OutputStream os, int maxLines) {
        super(os);
        this.maxLines = maxLines;
    }

    @Override
    public void write(int b) throws IOException {
        if (state != BODY || maxLines > 0) {
            super.write(b);
        }
        if (state == BODY) {
            if (b == '\n') {
                maxLines--;
            }
        } else if (state == START) {
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
