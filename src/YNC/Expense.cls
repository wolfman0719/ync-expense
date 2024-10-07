Class YNC.Expense Extends %Persistent
{

Property ReportMonth As %Integer;

Property ReportDay As %Integer;

Property PaymentTo As %String;

Property Accounts As %String;

Property Amount As %Integer;

Property Description As %String;

Property ISJReimbursement As %String;

Property OnBeHalf As %String;

Property ReportMD As %Integer [ Calculated, SqlComputeCode = {set {*} = $translate($justify({ReportMonth},2)," ",0)_$translate($justify({ReportDay},2)," ",0)}, SqlComputed ];

Storage Default
{
<Data name="ExpenseDefaultData">
<Value name="1">
<Value>%%CLASSNAME</Value>
</Value>
<Value name="2">
<Value>ReportMonth</Value>
</Value>
<Value name="3">
<Value>ReportDay</Value>
</Value>
<Value name="4">
<Value>PaymentTo</Value>
</Value>
<Value name="5">
<Value>Accounts</Value>
</Value>
<Value name="6">
<Value>Amount</Value>
</Value>
<Value name="7">
<Value>Description</Value>
</Value>
<Value name="8">
<Value>ISJReimbursement</Value>
</Value>
<Value name="9">
<Value>OnBeHalf</Value>
</Value>
</Data>
<DataLocation>^YNC.ExpenseD</DataLocation>
<DefaultData>ExpenseDefaultData</DefaultData>
<IdLocation>^YNC.ExpenseD</IdLocation>
<IndexLocation>^YNC.ExpenseI</IndexLocation>
<StreamLocation>^YNC.ExpenseS</StreamLocation>
<Type>%Storage.Persistent</Type>
}

}
