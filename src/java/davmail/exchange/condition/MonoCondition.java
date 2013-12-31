package davmail.exchange.condition;

import davmail.exchange.ExchangeSession;
import davmail.exchange.entity.Contact;

/**
 * Single search filter condition.
 */
public abstract class MonoCondition implements Condition {
    protected final String attributeName;
    protected final ExchangeSession.Operator operator;

    public MonoCondition(String attributeName, ExchangeSession.Operator operator) {
        this.attributeName = attributeName;
        this.operator = operator;
    }

    public boolean isEmpty() {
        return false;
    }

    public boolean isMatch(Contact contact) {
        String actualValue = contact.get(attributeName);
        return (operator == ExchangeSession.Operator.IsNull && actualValue == null) ||
                (operator == ExchangeSession.Operator.IsFalse && "false".equals(actualValue)) ||
                (operator == ExchangeSession.Operator.IsTrue && "true".equals(actualValue));
    }
}
