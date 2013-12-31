package davmail.exchange.dav;

import davmail.exchange.condition.Condition;

class DavNotCondition extends davmail.exchange.condition.NotCondition {
    DavNotCondition(Condition condition) {
        super(condition);
    }

    public void appendTo(StringBuilder buffer) {
        buffer.append("(Not ");
        condition.appendTo(buffer);
        buffer.append(')');
    }
}
