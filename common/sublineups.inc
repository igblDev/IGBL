<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConn, objRS

	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRS = Server.CreateObject("ADODB.RecordSet")   	
	Set objRSLineupsPics = Server.CreateObject("ADODB.RecordSet")

  objConn.Open Application("lineupstest_ConnectionString")
  Dim strDatabaseType

  strDatabaseType = "Access"
	ownerid = Request.querystring("ownerid")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
                  "Persist Security Info=False"

	ownerid = session("ownerid")	
	
  if ownerid = "" then
    GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	

  objRS.Open	"SELECT * FROM qrySubLineups WHERE qrySubLineups.ownerID =" & ownerid & " ", objConn,3,3,1


%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2016-2017</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/app.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
black{
	color:black;
}
red {
	color: red;
	font-size:10px;
}
td {
    vertical-align: middle;
		text-align: center;
		font-size:10px;
}
.panel-heading{

}
.alert-success {
    border-color: #468847;;
}
</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-success">
			<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Click the <strong><span class="glyphicon glyphicon-pencil"></span></strong> to edit line-ups.
			</div>
		</div>
	</div>
</div>
<%
   While Not objRS.EOF
%>
<% 
	if objRS.Fields("gameday").Value = date() and objRS.Fields("gameStaggerDeadline").Value < (time() - 1/24) then
		showPencil    = false
	else
		showPencil    = true 
	end if
	
	cFirstName    = ""
	cLastName     = ""
	sForward2     = ""
	f1LastName    = ""
	sGuard2       = ""
	
	sCenterBarps  = 0
	sForwardBarps = 0
	sForward2Barps= 0
	sGuardBarps   = 0
	sGuard2Barps  = 0

	sCenterTime   = 0
	sForwardTime  = 0
	sForward2Time = 0
	sGuardTime    = 0
	sGuard2Time   = 0
	
	sCenter       = objRS.Fields("sCenter").Value
	sForward      = objRS.Fields("sforward").Value
	sForward2     = objRS.Fields("sforward2").Value
	sGuard        = objRS.Fields("sguard").Value
	sGuard2       = objRS.Fields("sguard2").Value
	
	sCenterBarps  = objRS.Fields("sCenterBarps").Value
	sForwardBarps = objRS.Fields("sforwardBarps").Value
	sForward2Barps= objRS.Fields("sForward2Barps").Value
	sGuardBarps   = objRS.Fields("sguardBarps").Value
	sGuard2Barps  = objRS.Fields("sguard2Barps").Value

	sCenterTime   = objRS.Fields("sCenterTip").Value
	sForwardTime  = objRS.Fields("sForwardTip").Value
	sForward2Time = objRS.Fields("sForwardTip2").Value
	sGuardTime    = objRS.Fields("sGuardTip").Value
	sGuard2Time   = objRS.Fields("sGuardTip2").Value
					
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sCenter  & " = PID " , objConn,3,3,1
	ACenPic    = objRSLineupsPics.Fields("image").Value
	cFirstName = left(objRSLineupsPics.Fields("firstName").Value,1)
	cLastName  = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward  & " = PID " , objConn,3,3,1
	AForPic    = objRSLineupsPics.Fields("image").Value
	f1FirstName= left(objRSLineupsPics.Fields("firstName").Value,1)
	f1LastName = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward2  & " = PID " , objConn,3,3,1
	AFor2Pic   = objRSLineupsPics.Fields("image").Value
	f2FirstName= left(objRSLineupsPics.Fields("firstName").Value,1)
	f2LastName = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard  & " = PID " , objConn,3,3,1
	AGuaPic    = objRSLineupsPics.Fields("image").Value
	g1FirstName= left(objRSLineupsPics.Fields("firstName").Value,1)
	g1LastName = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard2 & " = PID " , objConn,3,3,1
	AGua2Pic   = objRSLineupsPics.Fields("image").Value
	g2FirstName= left(objRSLineupsPics.Fields("firstName").Value,1)
	g2LastName = left(objRSLineupsPics.Fields("lastName").Value,9)

	objRSLineupsPics.Close 
	
%>
<div class="container">
  <div class="row"> 
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-primary">
        <div class="panel-heading clearfix">
				<% if showPencil = true then %>
          <h4 class="panel-title">Date:&nbsp;<%= objRS.Fields("gameday").Value %><a href="lineups.asp?ownerid=<%= ownerid %>&Action=Retrieve Lineup&currentDate=<%= objRS.Fields("gameday").Value%>"><button type="submit" class="btn btn-xs btn-default pull-right"><span class="glyphicon glyphicon-pencil"></span></button></a></h4>
				<% else %>
          <h4 class="panel-title">Date:&nbsp;<%= objRS.Fields("gameday").Value %></h4>
				<%end if %>
        </div>
        <table class="table table-bordered table-striped table-condensed">
          <tr>
            <td style="background-color:#468847;color:white" width="20%"><strong>CEN</strong></td>
            <td style="background-color:#468847;color:white" width="20%"><strong>FOR</strong></td>
            <td style="background-color:#468847;color:white" width="20%"><strong>FOR</strong></td>
            <td style="background-color:#468847;color:white" width="20%"><strong>GUA</strong></td>
            <td style="background-color:#468847;color:white" width="20%"><strong>GUA</strong></td>
					</tr>
          <tr>
						<% if sCenter 	> 9900 then %>
						<td><strong><red><%=cLastName%></red></strong></td>
						<% else %>
						<td><black><%=cLastName%></black></td>
						<% end if %>
						
						<% if sForward 	> 9900 then %>
						<td><strong><red><%=f1LastName%></red></strong></td>
						<% else %>
						<td><black><%=f1LastName%></black></td>
						<% end if %>
						
						<% if sForward2 > 9900 then %>
						<td><strong><red><%=f2LastName%></red></strong></td>
						<% else %>
						<td><black><%=f2LastName%></black></td>
						<% end if %>
						
						<% if sGuard 		> 9900 then %>
						<td><strong><red><%=g1LastName%></red></strong></td>
						<% else %>
						<td><black><%=g1LastName%></black></td>
						<% end if %>
						
						<% if sGuard2 	> 9900 then %>
						<td><strong><red><%=g2LastName%></red></strong></td>
						<% else %>
						<td><black><%=g2LastName%></black></td>
						<% end if %>
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