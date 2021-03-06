package davmail.exchange.condition;

import davmail.exchange.entity.Contact;

/**
 * Exchange search filter.
 */
public interface Condition {
    /**
     * Append condition to buffer.
     *
     * @param buffer search filter buffer
     */
    void appendTo(StringBuilder buffer);

    /**
     * True if condition is empty.
     *
     * @return true if condition is empty
     */
    boolean isEmpty();

    /**
     * Test if the contact matches current condition.
     *
     * @param contact Exchange Contact
     * @return true if contact matches condition
     */
    boolean isMatch(Contact contact);
}
