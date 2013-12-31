package davmail.exchange.dav;

import davmail.exchange.ExchangeSession;
import davmail.exchange.condition.Condition;

class DavMultiCondition extends davmail.exchange.condition.MultiCondition {
    DavMultiCondition(ExchangeSession.Operator operator, Condition... condition) {
        super(operator, condition);
    }

    public void appendTo(StringBuilder buffer) {
        boolean first = true;

        for (Condition condition : conditions) {
            if (condition != null && !condition.isEmpty()) {
                if (first) {
                    buffer.append('(');
                    first = false;
                } else {
                    buffer.append(' ').append(operator).append(' ');
                }
                condition.appendTo(buffer);
            }
        }
        // at least one non empty condition
        if (!first) {
            buffer.append(')');
        }
    }
}
