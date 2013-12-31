package davmail.exchange.condition;

import davmail.exchange.entity.Contact;

/**
 * Not condition.
 */
public abstract class NotCondition implements Condition {
    protected final Condition condition;

    public NotCondition(Condition condition) {
        this.condition = condition;
    }

    public boolean isEmpty() {
        return condition.isEmpty();
    }

    public boolean isMatch(Contact contact) {
        return !condition.isMatch(contact);
    }
}
