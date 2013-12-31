package davmail.exchange.condition;

import davmail.exchange.ExchangeSession;
import davmail.exchange.entity.Contact;

import java.util.ArrayList;
import java.util.List;

/**
 * Multiple condition.
 */
public abstract class MultiCondition implements Condition {
    protected final ExchangeSession.Operator operator;
    protected final List<Condition> conditions;

    public MultiCondition(ExchangeSession.Operator operator, Condition... conditions) {
        this.operator = operator;
        this.conditions = new ArrayList<Condition>();
        for (Condition condition : conditions) {
            if (condition != null) {
                this.conditions.add(condition);
            }
        }
    }

    /**
     * Conditions list.
     *
     * @return conditions
     */
    public List<Condition> getConditions() {
        return conditions;
    }

    /**
     * Condition operator.
     *
     * @return operator
     */
    public ExchangeSession.Operator getOperator() {
        return operator;
    }

    /**
     * Add a new condition.
     *
     * @param condition single condition
     */
    public void add(Condition condition) {
        if (condition != null) {
            conditions.add(condition);
        }
    }

    public boolean isEmpty() {
        boolean isEmpty = true;
        for (Condition condition : conditions) {
            if (!condition.isEmpty()) {
                isEmpty = false;
                break;
            }
        }
        return isEmpty;
    }

    public boolean isMatch(Contact contact) {
        if (operator == ExchangeSession.Operator.And) {
            for (Condition condition : conditions) {
                if (!condition.isMatch(contact)) {
                    return false;
                }
            }
            return true;
        } else if (operator == ExchangeSession.Operator.Or) {
            for (Condition condition : conditions) {
                if (condition.isMatch(contact)) {
                    return true;
                }
            }
            return false;
        } else {
            return false;
        }
    }

}
