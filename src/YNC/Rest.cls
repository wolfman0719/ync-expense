Class YNC.Rest Extends %CSP.REST
{

Parameter CONVERTINPUTSTREAM = 1;

Parameter HandleCorsRequest = 1;

XData UrlMap
{
<Routes>
  <Route Url="/SaveExpenseItem" Method="POST" Call="SaveExpenseItem"/>
  <Route Url="/SearchExpenseItem/:keyword" Method="GET" Call="SearchExpenseItem"/>
  <Route Url="/ExpenseItemGetById/:id" Method="GET" Call="ExpenseItemGetById"/>
  <Route Url="/DeleteExpenseItemById/:id" Method="GET" Call="DeleteExpenseItemById"/>
</Routes>
}

ClassMethod ExpenseItemGetById(pId As %Integer) As %Status
{
  set status = $$$OK

  try {
	
    if $data(%request) {
      set %response.ContentType="application/json"
      set %response.CharSet = "utf-8"
    }
    set expenseitem = ##class(YNC.ExpenseItem).%OpenId(pId,,.status)	
    $$$ThrowOnError(status)  
    set return = {}
    set return.id = pId
    set return.paymentto = expenseitem.PaymentTo
    set return.description = expenseitem.Description
    set return.amount = expenseitem.Amount
    set return.onbehalf = expenseitem.OnBeHalf
    set return.accounts = expenseitem.Accounts
    write return.%ToJSON()
	  	  
  }
  catch e {
		
    set status = e.AsStatus()
		
  }
  
  quit status
}

ClassMethod DeleteExpenseItemById(pId As %Integer) As %Status
{
  set status = $$$OK

  try {
	
    if $data(%request) {
      set %response.ContentType="application/json"
      set %response.CharSet = "utf-8"
    }
    set status = ##class(YNC.ExpenseItem).%DeleteId(pId)
    $$$ThrowOnError(status)  
	  	  
  }
  catch e {
		
    set status = e.AsStatus()
		
  }
  
  quit status
}

ClassMethod SaveExpenseItem() As %Status
{
  set status = $$$OK
  try {
	
    if $data(%request) {
      set %response.ContentType="application/json"
      set %response.CharSet = "utf-8"
    }

    set len = 10000
	set jsontext = %request.Content.Read(.len,.status)
	set json = {}.%FromJSON(jsontext)

    set new = json.new

    if new {

      set expenseitem = ##class(YNC.ExpenseItem).%New()

    }
    else {

      set id=$translate(json.id,$zzenkaku("0123456789"),"0123456789")
      set expenseitem = ##class(YNC.ExpenseItem).%OpenId(id,,.status)	
      $$$ThrowOnError(status)  

    }

    set expenseitem.Accounts =json.accounts
    set expenseitem.Amount =json.amount
    set expenseitem.Description = json.description
    set expenseitem.OnBeHalf = json.onbehalf
    set expenseitem.PaymentTo = json.paymentto
    set status = expenseitem.%Save()
    $$$ThrowOnError(status) 

  }
  catch e {
		
    set status = e.AsStatus()
		
  }
  quit status
}

ClassMethod SearchExpenseItem(pKeyword As %String) As %Status
{
  set status = $$$OK
  
  try  {
	
    if $data(%request) {
      set %response.ContentType="application/json"
      set %response.CharSet = "utf-8"
    }

    set pKeyword = $extract(pKeyword,2,*)

  if pKeyword = "" {	  
    set sql = "SELECT id,description from ync.expenseitem"
  }
  else {
    set sql = "SELECT id,description from ync.expenseitem where description %startswith(?)"
  }
    
  set statemt=##class(%SQL.Statement).%New()
  set status=statemt.%Prepare(sql)
  if sql [ "?" {
    set rset=statemt.%Execute(pKeyword)
  }
  else {
    set rset=statemt.%Execute()
  }
  set return = []
  while rset.%Next()  {
    set robj = {}
    set robj.id = rset.%Get("Id")
    set robj.description = rset.%Get("description")
    do return.%Push(robj)
		    
  }
	   
  write return.%ToJSON()
	  	  
  }
  catch e {
		
    set status = e.AsStatus()
		
  }
  
  quit status
}

ClassMethod OnHandleCorsRequest() As %Status
{
  // Get the origin
  //Set tOrigin=$Get(%request.CgiEnvs("HTTP_ORIGIN"))
  Set tOrigin="*"

  // Allow requested origin
  Do ..SetResponseHeaderIfEmpty("Access-Control-Allow-Origin",tOrigin)

  set tMethod = "GET,POST,PUT,DELETE,OPTIONS, PATCH"
  Do ..SetResponseHeaderIfEmpty("Access-Control-Allow-Method",tMethod)

  set tHeader = "Content-Type,Authorization,X-Requested-With"
  Do ..SetResponseHeaderIfEmpty("Access-Control-Allow-Headers",tHeader)

  Do ..SetResponseHeaderIfEmpty("Access-Control-Allow-Credentials","true")

  quit $$$OK
}

}
