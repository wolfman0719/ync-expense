Class YNC.ExpenseItem Extends %Persistent
{

Property Description As %String;

Property PaymentTo As %String;

Property Accounts As %String;

Property Amount As %Integer;

Property OnBeHalf As %String;

/// Description
Index DescriptionIndex On Description [ IdKey ];

}
