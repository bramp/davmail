package davmail.exchange.dav;

import davmail.exchange.ExchangeSession;
import davmail.exchange.entity.Contact;
import davmail.util.StringUtil;

class DavAttributeCondition extends davmail.exchange.condition.AttributeCondition {
    protected boolean isIntValue;

    DavAttributeCondition(String attributeName, ExchangeSession.Operator operator, String value) {
        super(attributeName, operator, value);
    }

    DavAttributeCondition(String attributeName, ExchangeSession.Operator operator, int value) {
        super(attributeName, operator, String.valueOf(value));
        isIntValue = true;
    }

    public void appendTo(StringBuilder buffer) {
        Field field = Field.get(attributeName);
        buffer.append('"').append(field.getUri()).append('"');
        buffer.append(DavExchangeSession.OPERATOR_MAP.get(operator));
        //noinspection VariableNotUsedInsideIf
        if (field.cast != null) {
            buffer.append("CAST (\"");
        } else if (!isIntValue && !field.isIntValue()) {
            buffer.append('\'');
        }
        if (ExchangeSession.Operator.Like == operator) {
            buffer.append('%');
        }
        if ("urlcompname".equals(field.alias)) {
            buffer.append(StringUtil.encodeUrlcompname(StringUtil.davSearchEncode(value)));
        } else if (field.isIntValue()) {
            // check value
            try {
                Integer.parseInt(value);
                buffer.append(value);
            } catch (NumberFormatException e) {
                // invalid value, replace with 0
                buffer.append('0');
            }
        } else {
            buffer.append(StringUtil.davSearchEncode(value));
        }
        if (ExchangeSession.Operator.Like == operator || ExchangeSession.Operator.StartsWith == operator) {
            buffer.append('%');
        }
        if (field.cast != null) {
            buffer.append("\" as '").append(field.cast).append("')");
        } else if (!isIntValue && !field.isIntValue()) {
            buffer.append('\'');
        }
    }

    public boolean isMatch(Contact contact) {
        String lowerCaseValue = value.toLowerCase();
        String actualValue = contact.get(attributeName);
        ExchangeSession.Operator actualOperator = operator;
        // patch for iCal or Lightning search without galLookup
        if (actualValue == null && ("givenName".equals(attributeName) || "sn".equals(attributeName))) {
            actualValue = contact.get("cn");
            actualOperator = ExchangeSession.Operator.Like;
        }
        if (actualValue == null) {
            return false;
        }
        actualValue = actualValue.toLowerCase();
        return (actualOperator == ExchangeSession.Operator.IsEqualTo && actualValue.equals(lowerCaseValue)) ||
                (actualOperator == ExchangeSession.Operator.Like && actualValue.contains(lowerCaseValue)) ||
                (actualOperator == ExchangeSession.Operator.StartsWith && actualValue.startsWith(lowerCaseValue));
    }
}
