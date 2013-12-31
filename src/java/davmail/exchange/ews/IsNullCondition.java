package davmail.exchange.ews;

import davmail.exchange.condition.Condition;
import davmail.exchange.entity.Contact;

class IsNullCondition implements Condition, SearchExpression {
    protected final String attributeName;

    IsNullCondition(String attributeName) {
        this.attributeName = attributeName;
    }

    public void appendTo(StringBuilder buffer) {
        buffer.append("<t:Not><t:Exists>");
        Field.get(attributeName).appendTo(buffer);
        buffer.append("</t:Exists></t:Not>");
    }

    public boolean isEmpty() {
        return false;
    }

    public boolean isMatch(Contact contact) {
        String actualValue = contact.get(attributeName);
        return actualValue == null;
    }

}
