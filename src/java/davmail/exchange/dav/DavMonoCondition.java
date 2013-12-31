package davmail.exchange.dav;

import davmail.exchange.ExchangeSession;

class DavMonoCondition extends davmail.exchange.condition.MonoCondition {
    DavMonoCondition(String attributeName, ExchangeSession.Operator operator) {
        super(attributeName, operator);
    }

    public void appendTo(StringBuilder buffer) {
        buffer.append('"').append(Field.get(attributeName).getUri()).append('"');
        buffer.append(DavExchangeSession.OPERATOR_MAP.get(operator));
    }
}
