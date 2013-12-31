package davmail.imap;

import davmail.util.StringUtil;

import java.util.StringTokenizer;

class IMAPTokenizer extends StringTokenizer {
    IMAPTokenizer(String value) {
        super(value);
    }

    @Override
    public String nextToken() {
        return StringUtil.removeQuotes(nextQuotedToken());
    }

    public String nextQuotedToken() {
        StringBuilder nextToken = new StringBuilder();
        nextToken.append(super.nextToken());
        while (hasMoreTokens() && nextToken.length() > 0 && nextToken.charAt(0) == '"'
                && (nextToken.charAt(nextToken.length() - 1) != '"' || nextToken.length() == 1)) {
            nextToken.append(' ').append(super.nextToken());
        }
        while (hasMoreTokens() && nextToken.length() > 0 && nextToken.charAt(0) == '('
                && nextToken.charAt(nextToken.length() - 1) != ')') {
            nextToken.append(' ').append(super.nextToken());
        }
        while (hasMoreTokens() && nextToken.length() > 0 && nextToken.indexOf("[") != -1
                && nextToken.charAt(nextToken.length() - 1) != ']') {
            nextToken.append(' ').append(super.nextToken());
        }
        return nextToken.toString();
    }
}
