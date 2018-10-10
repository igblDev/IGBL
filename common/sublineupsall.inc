<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""


	Set objConn         = Server.CreateObject("ADODB.Connection")
	Set objRS           = Server.CreateObject("ADODB.RecordSet")    
	Set objRSLineupsPics= Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	
	strDatabaseType = "Access"
	ownerid = Request.querystring("ownerid")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	ownerid = Request.querystring("ownerid")
	objRS.Open	"SELECT * FROM qrySubLineups", objConn,3,3,1
	
	ownerid = session("ownerid")	
	
	if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	


%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2016-2017</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/app.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
</head>
<style>
black{
	color:black;
}
red {
	color: red;
	font-size:10px;
}
td {
    vertical-align:middle;
		text-align:center;
		font-size:10px;
}
</style>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<%
   While Not objRS.EOF
%>
<% 
	sCenter       = objRS.Fields("sCenter").Value
	sForward      = objRS.Fields("sforward").Value
	sForward2     = objRS.Fields("sforward2").Value
	sGuard        = objRS.Fields("sguard").Value
	sGuard2       = objRS.Fields("sguard2").Value
	
	sCenterBarps  = objRS.Fields("sCenterBarps").Value
	sForwardBarps = objRS.Fields("sforwardBarps").Value
	sForward2Barps= objRS.Fields("sforward2Barps").Value
	sGuardBarps   = objRS.Fields("sguardBarps").Value
	sGuard2Barps  = objRS.Fields("sguard2Barps").Value

	sCenterTime   = objRS.Fields("sCenterTip").Value
	sForwardTime  = objRS.Fields("sForwardTip").Value
	sForward2Time = objRS.Fields("sForwardTip2").Value
	sGuardTime    = objRS.Fields("sGuardTip").Value
	sGuard2Time   = objRS.Fields("sGuardTip2").Value
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sCenter  & " = PID " , objConn,3,3,1
	ACenPic       = objRSLineupsPics.Fields("image").Value
	'cFirstName    = left(objRSLineupsPics.Fields("firstName").Value,1)
	cLastName     = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward  & " = PID " , objConn,3,3,1
	AForPic       = objRSLineupsPics.Fields("image").Value
	'f1FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	f1LastName    = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward2  & " = PID " , objConn,3,3,1
	AFor2Pic      = objRSLineupsPics.Fields("image").Value
	'f2FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	f2LastName    = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard  & " = PID " , objConn,3,3,1
	AGuaPic       = objRSLineupsPics.Fields("image").Value
	'g1FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	g1LastName    = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard2 & " = PID " , objConn,3,3,1
	AGua2Pic      = objRSLineupsPics.Fields("image").Value
	'g2FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	g2LastName    = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close 
%>
<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-primary">
        <div class="panel-heading clearfix">
          <h4 class="panel-title"><%= objRS.Fields("teamName").Value %></h4>
        </div>
        <table class="table table-bordered table-striped table-condensed">
          <tr>
            <td align="center" colspan="5" style="background-color:white;color:black"><strong>Game Date:&nbsp;<%= objRS.Fields("gameday").Value %></strong></td>
          </tr>
          <tr>
            <td style="background-color:#468847;color:white;text-align:center" width="20%"><strong>CEN</strong></td>
            <td style="background-color:#468847;color:white;text-align:center" width="20%"><strong>FOR</strong></td>
            <td style="background-color:#468847;color:white;text-align:center" width="20%"><strong>FOR</strong></td>
            <td style="background-color:#468847;color:white;text-align:center" width="20%"><strong>GUA</strong></td>
            <td style="background-color:#468847;color:white;text-align:center" width="20%"><strong>GUA</strong></td>
					</tr>
          <tr>
            <td><black><%=cLastName%></black></td>
            <td><black><%=f1LastName%></black></td>
            <td><black><%=f2LastName%></black></td>
            <td><black><%=g1LastName%></black></td>
            <td><black><%=g2LastName%></black></td>
          </tr>
          <tr>
            <td><%=sCenterTime%></td>
            <td><%=sForwardTime%></td>
            <td><%=sForward2Time%></td>
            <td><%=sGuardTime%></td>
            <td><%=sGuard2Time%></td>
          </tr>  
	 </table>
      </div>
    </div>
  </div>
</div>
<%
   objRS.MoveNext
   Wend
%>
<%
  objRS.Close
  ObjConn.Close

  Set objRS = Nothing
  objRSLineupsPics = Nothing
  Set objConn = Nothing

  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>