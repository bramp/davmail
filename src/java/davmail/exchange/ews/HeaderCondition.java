package davmail.exchange.ews;

import davmail.exchange.ExchangeSession;

class HeaderCondition extends EwsExchangeSession.AttributeCondition {

    HeaderCondition(String attributeName, String value) {
        super(attributeName, ExchangeSession.Operator.Contains, value);
        containmentMode = ContainmentMode.Substring;
        containmentComparison = ContainmentComparison.IgnoreCase;
    }

    @Override
    protected FieldURI getFieldURI() {
        return new ExtendedFieldURI(ExtendedFieldURI.DistinguishedPropertySetType.InternetHeaders, attributeName);
    }

}
