package davmail.exchange.condition;

import davmail.exchange.ExchangeSession;

/**
 * Attribute condition.
 */
public abstract class AttributeCondition implements Condition {
    protected final String attributeName;
    protected final ExchangeSession.Operator operator;
    protected final String value;

    public AttributeCondition(String attributeName, ExchangeSession.Operator operator, String value) {
        this.attributeName = attributeName;
        this.operator = operator;
        this.value = value;
    }

    public boolean isEmpty() {
        return false;
    }

    /**
     * Get attribute name.
     *
     * @return attribute name
     */
    public String getAttributeName() {
        return attributeName;
    }

    /**
     * Condition value.
     *
     * @return value
     */
    public String getValue() {
        return value;
    }

}
