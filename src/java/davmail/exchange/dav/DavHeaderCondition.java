package davmail.exchange.dav;

import davmail.exchange.ExchangeSession;

class DavHeaderCondition extends DavAttributeCondition {

    DavHeaderCondition(String attributeName, ExchangeSession.Operator operator, String value) {
        super(attributeName, operator, value);
    }

    @Override
    public void appendTo(StringBuilder buffer) {
        buffer.append('"').append(Field.getHeader(attributeName).getUri()).append('"');
        buffer.append(DavExchangeSession.OPERATOR_MAP.get(operator));
        buffer.append('\'');
        if (ExchangeSession.Operator.Like == operator) {
            buffer.append('%');
        }
        buffer.append(value);
        if (ExchangeSession.Operator.Like == operator) {
            buffer.append('%');
        }
        buffer.append('\'');
    }
}
