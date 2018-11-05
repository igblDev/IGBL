<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	

	Dim strSQL,objConn,objRSParms,objRSParmsAmts

	GetAnyParameter "Action", sAction	

   	
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"
		
%>	
	<!--#include virtual="Common/session.inc"-->	
<%
	Set objRSParms = Server.CreateObject("ADODB.RecordSet")
	Set objRSParmsAmts = Server.CreateObject("ADODB.RecordSet")
	

	'Response.Write "Action Contains = "&sAction&" <br> "	
	'Response.Write "New Date = "&CalDate&" <br> "


	select case sAction
		
	  case ""
			objRSParmsAmts.Open "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT >= 0 ", objConn
			objRSParms.Open     "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT IS NULL AND PARAM_NAME <> 'STAGGER_WINDOW' ", objConn
		
		case "Update DateParms"
			splitParms= Split(Request.Form("chkBoxParms"),";")
			paramName = splitParms(0)
			paramDate = splitParms(1)
			CalDate   = Request.Form("newDate")
			NewDate   = mid(Request.Form("newDate"), Instr(Request.Form("newDate"),"2"), 10)

			'Response.Write " Parameter Name = "&paramName&" <br> "
			'Response.Write " Parameter Old Parm Date = "&paramDate&" <br> "
			'Response.Write " Parameter New Parm Date = "&CalDate&" <br> "
			'Response.Write " Parameter New Parm Date = "&NewDate&" <br> "
			
			strSQL = "update tblParameterCtl set param_date = #"& NewDate &"# where param_name = '"& paramName & "'"
			objConn.Execute strSQL
			'Response.Write "Sql = " & strSQL  & "<br>" 
						
			if paramName = "TRADE_DEADLINE" and (cDate(NewDate) >= date()) then
			   strSQL = "update tblOwners set acceptTradeOffers = 1" 
               objConn.Execute strSQL			   
			end if

			sAction = ""
			objRSParmsAmts.Open "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT >= 0 ", objConn
			objRSParms.Open     "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT IS NULL AND PARAM_NAME <> 'STAGGER_WINDOW' ", objConn
		
		case "Update DollarParms"		
			splitParms2 = Split(Request.Form("chkBox2Parms"),";")
			paramName2  = splitParms2(0)
			paramAmount2= splitParms2(1)   
			newFees     = Split(Request.Form("newFeeAmt"),",")			
			newFeesCnt  = Request.Form("newFeeAmt").count
			
			For i = 0 To newFeesCnt
				'response.write("The Fee Amount is " & newFees(i) & "<br />")
				if newFees(i)= " " or IsNull(newFees(i)) then
					'Do Nothing
				else 	
					newFeeAmt =  newFees(i)
				end if	
			Next
			
			'Response.Write " Parameter Name = "&paramName2&" <br> "
			'Response.Write " Parameter Old Fee Amt = "&paramAmount2&" <br> "
			'Response.Write " Parameter New Fee Amt = "&newFeeAmt&" <br> "  
			'Response.Write " Count of Values = "&newFeesCnt&" <br> "  
			 
			strSQL = "update tblParameterCtl set param_amount = '"& newFeeAmt &"' where param_name = '"& paramName2 & "'"  
			'Response.Write "Sql = " & strSQL  & "<br>" 
			objConn.Execute strSQL

			sAction = ""
			objRSParmsAmts.Open "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT >= 0 ", objConn
			objRSParms.Open     "SELECT * FROM tblParameterCtl WHERE PARAM_AMOUNT IS NULL AND PARAM_NAME <> 'STAGGER_WINDOW' ", objConn
			
	end select	

%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2017-2018</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>

</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Configure Parameters</strong>
			</div>
		</div>
	</div>
</div>
<%
	if sAction = "" then
%>
<form action="parmMaint.asp" name="parmMaint" id="Player" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_pid" value="<%=pid%>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#dates">Dates</a></li>
				<li><a data-toggle="tab" href="#dollars">Dollars</a></li>
			</ul>
		</div>
	</div>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="dates" class="tab-pane fade in active">
					<div class="row">
						<div class="col-md-12 col-sm-12 col-xs-12">
							<div class="panel panel-override">
								<div class="panel-body">
									<table class="table table-custom-black table-responsive table-bordered table-condensed">
										<tr>
											<th style="width:45%;">Current Date</th>
											<th style="width:45%;">New Date</th>
											<th style="width:10%;"></th>
										</tr>
										<%
										While Not objRSParms.EOF
										%>
										<tr style="background-color:white;">
											<td><input type="text" class="form-control" name="param_date" readonly value="<%=objRSParms.Fields("param_date")%>"id="policyIDId"></td>
											<td><input type="date" class="form-control input-xs" name="newDate"  id="persistentKeyId"></td>
											<td style="width:10%;text-align:center;vertical-align:middle;"><input type="checkbox" value="<%=objRSParms.Fields("param_name").Value & ";" & objRSParms.Fields("param_date").Value & ";"%>" name="chkBoxParms"></td>
										</tr>
										<tr style="background-color:white;">
											<td colspan="3" ><button type="submit" value="Update DateParms" name="Action" class="btn btn-sm btn-block btn-default"><%=objRSParms.Fields("param_name")%></button></td>
										</tr>
										<%
										objRSParms.MoveNext
										Wend
										%>
									</table>
								</div>
							</div>
						</div>
					</div>
				</div>
				<div id="dollars" class="tab-pane fade">
					<div class="row">
						<div class="col-md-12 col-sm-12 col-xs-12">
							<div class="panel panel-override">
								<div class="panel-body">
									<table class="table table-custom-black table-responsive table-bordered table-condensed">
										<tr>
											<th style="width:5%;"></th>
											<th style="width:35%;">Old</th>
											<th style="width:35%;">New</th>
											<th style="width:25%;">Parameter</th>
										</tr>
										<%
										While Not objRSParmsAmts.EOF
										%>
										<tr style="background-color:white;">
											<td style="width:10%;text-align:center;vertical-align:middle;"><input type="checkbox" value="<%=objRSParmsAmts.Fields("param_name").Value & ";" & objRSParmsAmts.Fields("param_amount").Value & ";"%>" name="chkBox2Parms"></td>
											<td><input type="text" class="form-control" name="param_amount" readonly value="<%=objRSParmsAmts.Fields("param_amount")%>"id="policyIDId"></td>
											<td><input type="number" class="form-control input-xs" name="newFeeAmt"  id="persistentKeyId"></td>
											<td colspan="3" ><button type="submit" value="Update DollarParms" name="Action" class="btn btn-sm btn-block btn-default"><%=objRSParmsAmts.Fields("param_name")%></button></td>

										</tr>
										<!--<tr style="background-color:white;">
											<td colspan="3" ><button type="submit" value="Update DollarParms" name="Action" class="btn btn-sm btn-block btn-default"><redText><strong><%=objRSParmsAmts.Fields("param_name")%><redText></button></td>
										</tr>-->	
										<%
										objRSParmsAmts.MoveNext
										Wend
										%>					
									</table>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
</form>
<%end if %>

<%  
	objRSParms.Close
	objRSParmsAmts.Close
	objRSPStatus.Close
  Set objRSParms = Nothing
  ObjConn.Close
  Set objConn = Nothing
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>