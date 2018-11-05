<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
							 "Data Source=lineupstest.mdb;" & _
							 "Persist Security Info=False"

	dim	addPlayerCnt,releasePlayerCnt,objRSWork, objRSobjRS
	
	Set objRSWork  = Server.CreateObject("ADODB.RecordSet") 
	Set objRS      = Server.CreateObject("ADODB.RecordSet")
		
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
	wPickUp = objRSWork.Fields("param_amount").Value
	objRSWork.Close

	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE' ",objConn
	wTrade = objRSWork.Fields("param_amount").Value
	objRSWork.Close
	
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT' ",objConn
	wRental = objRSWork.Fields("param_amount").Value
	objRSWork.Close
	
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT_PENALTY' ",objConn
	wRentPenalty = objRSWork.Fields("param_amount").Value
	objRSWork.Close

%>
	<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
black{
	color:black;
}
span.reject {
	color: #9a1400;
	font-weight: bold;
}
.item {
	background: #333;
	text-align: center;
	height: 120px !important;}
	table.dataTable tbody th, 
	table.dataTable tbody td {
	padding: 1px 5px;
}
span.greenIcon{
	color:green;
	font-weight: 500;
}
span.blackIcon{
	color:black;
	font-weight: 500;
}
span.blueTitle {
	color:#01579B;
	font-weight: bold;
	text-transform: uppercase;
}
span.redIcon{
	color:#9a1400;
	font-weight: 500;
}
#load{
	width:100%;
	height:100%;
	position:fixed;
	z-index:9999;
	background:url("https://www.creditmutuel.fr/cmne/fr/banques/webservices/nswr/images/loading.gif") no-repeat center center rgba(0,0,0,0.25)
}
select {
    color: black;
    background-color: yellowgreen;;
}
span.dollar {
	color: black;
	font-weight: 500;
	text-transform: uppercase;
}	
.mark, mark {
    padding: .2em;
    background-color: yellow;
}
.nav-tabs {
    border-bottom: 2px solid black;
}
</style>
</head>
<body>
<div id="load"></div>
<script language="JavaScript" type="text/javascript">
$(document).ready(function() {
    $('#example').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false,
				"lengthMenu": [[15, 30, 50, -1], [15, 30, 50, "All"]]
    } );
} );
$(document).ready(function() {
    $('#example2').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false,
				"lengthMenu": [[15, 30, 50, -1], [15, 30, 50, "All"]]
    } );
} );
$(document).ready(function() {
    $('#example3').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false,
				"lengthMenu": [[15, 30, 50, -1], [15, 30, 50, "All"]]
    } );
} );
document.onreadystatechange = function () {
  var state = document.readyState
  if (state == 'interactive') {
       document.getElementById('contents').style.visibility="hidden";
  } else if (state == 'complete') {
      setTimeout(function(){
         document.getElementById('interactive');
         document.getElementById('load').style.visibility="hidden";
         document.getElementById('contents').style.visibility="visible";
      },1000);
  }
}
</script>


<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>League Finances</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#league"><i class="fas fa-users"></i>&nbsp;Activity</a></li>
				<li><a data-toggle="tab" href="#waivers"><i class="far fa-money-bill-alt"></i>&nbsp;Waivers</a></li>
				<li><a data-toggle="tab" href="#mine"><i class="far fa-sort-numeric-down"></i>&nbsp;Waiver Order</a></li>
				<li><a data-toggle="tab" href="#finance"><i class="fas fa-usd-circle fa-lg"></i>&nbsp;MISC</a></li>
				<!--<li><a data-toggle="tab" href="#txncosts">Fees</a></li>-->
			</ul>
		</div>
	</div>
	</br>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="league"  class="tab-pane fade in active">
				<% 				
				objRS.Open "Select * from qryTransactions where waiverbid is null ",objConn 	

				%>

						<table class="table table-custom-black table-responsive table-bordered table-condensed" width="100%" class="display" id="example3">
							<thead>
								<tr>
									<th style="width:55%;">Team | Date</th>
									<th style="width:45%;">Player</th>
								</tr>
							</thead>
							<tbody>
							<%
							While Not objRS.EOF
							addPlayerCnt = 1
							releasePlayerCnt = 1
							wTeamName = replace(UCase(objRS.Fields("shortname").Value), "THE ", "")
							'Response.Write "wTeamName  = " & wTeamName & ".<br>"
							%>
							<tr bgcolor="#FFFFFF">		
							<%if objRS.Fields("ownerID").Value = ownerid then%>
								<td class="success big" style="text-align:left;vertical-align:top;"><span class="blueTitle"><strong><%=wTeamName%></strong></span></br><%=objRS.Fields("TransDate").value%></br>
								<% if objRS.Fields("TransType").value = "Rejected " or objRS.Fields("TransType").value  = "Rental Not Played " or objRS.Fields("TransType").value = "Rejected (Roster Full) " or objRS.Fields("TransType").value = "Rejected (Bid > Balance) "  then %>
									<strong><span class="reject"><%=objRS.Fields("TransType").value%></span></strong>
								<%elseif objRS.Fields("TransType").value = "Rental " then %>
									<strong><span class="orange"><%=objRS.Fields("TransType").value%></span></strong>	
								<%elseif objRS.Fields("TransType").value = "Added " then %>
									<strong><span class="auctionText"><%=objRS.Fields("TransType").value%></span></strong>													
								<%else%>
									<strong><span class="greenTrade"><%=objRS.Fields("TransType").value%></span></strong>								
								<%end if%>
								&nbsp;<span class="dollar"><%= FormatCurrency(objRS.Fields("TransCost").value)%></span>
								
								<%if objRS.Fields("waiverbid").Value > 0 then %>
								   <%if objRS.Fields("TransType").Value = "Added "  then%>
								      </br><mark><span class="dollar">Winning Bid</span>&nbsp;<span class="dollar">- [<%=objRS.Fields("waiverbid").value%>]</span><mark>
								   <%else%>
								      </br><span class="gameTip">Bid</span>&nbsp;<span class="dollar">- [<%=objRS.Fields("waiverbid").value%>]</span>
								   <%end if%>	
								<%end if%>	
								</td>
							<%else%>
							
								<td class="big" style="text-align:left"><span class="blueTitle"><%=wTeamName%></span></br><%=objRS.Fields("TransDate").value%></br>
								<% if objRS.Fields("TransType").value = "Rejected " or objRS.Fields("TransType").value  = "Rental Not Played " or objRS.Fields("TransType").value = "Rejected (Roster Full) " or objRS.Fields("TransType").value = "Rejected (Bid > Balance) "  then %>
									<strong><span class="reject"><%=objRS.Fields("TransType").value%></span></strong>
								<%elseif objRS.Fields("TransType").value = "Rental " then %>
									<strong><span class="orange"><%=objRS.Fields("TransType").value%></span></strong>	
								<%elseif objRS.Fields("TransType").value = "Added " then %>
									<strong><span class="auctionText"><%=objRS.Fields("TransType").value%></span></strong>													
								<%else%>
										<strong><span class="greenTrade"><%=objRS.Fields("TransType").value%></span></strong>																	
								<%end if%>
								&nbsp;<span class="dollar"><%=FormatCurrency(objRS.Fields("TransCost").value)%></span>
								
								<%if objRS.Fields("waiverbid").Value > 0 then%>
								   <%if objRS.Fields("TransType").Value = "Added "  then%>
								      </br><mark><span class="dollar">Winning Bid</span>&nbsp;<span class="dollar">- [<%=objRS.Fields("waiverbid").value%>]</span></mark>
								   <%else%>
								      </br><span class="gameTip">Bid</span>&nbsp;<span class="dollar">- [<%=objRS.Fields("waiverbid").value%>]</span>
								   <%end if%>
								<%end if%>									
								</td>
							<%end if%>	
								<td>
									<table>
									<%
										While cint(addPlayerCnt) <= cint(objRS.Fields("transAddPlayerCnt").value)
									%>
										<%if cint(addPlayerCnt) = 1 then %>
											<%wRetcd = GetName(objRS.Fields("transAddPlayer1").Value,wFName,wLName)
												wTransType = left(objRS.Fields("TransType").value,6)
												'Response.Write "Transaction Type  = " & wTransType & ".<br>"
											%>
											<%if wTransType = "Waiver" or wTransType = "Reject" then %>
												<tr><td class="big" style="text-align:center;width:4%;"><span class="blackIcon"><i class="fa fa-lock" aria-hidden="true"></i></span></td><td class="big" style="text-align:left;text-decoration: line-through;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
											<%else%>
												<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
											<%end if%>
										<%elseif cint(addPlayerCnt) = 2 then %>
											<%wRetcd = GetName(objRS.Fields("transAddPlayer2").Value,wFName,wLName)%>
											<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif cint(addPlayerCnt) = 3 then%>
												<%wRetcd = GetName(objRS.Fields("transAddPlayer3").Value,wFName,wLName)%>
											<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%end if%>
									<%
										addPlayerCnt = addPlayerCnt +1
										Wend
									%>
									<%
										While releasePlayerCnt <= objRS.Fields("transReleasePlayerCnt").value
									%>								
										<%if releasePlayerCnt = 1 then %>	
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer1").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif releasePlayerCnt = 2 then %>
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer2").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif releasePlayerCnt = 3 then%>
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer3").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%end if%>
									<%
										releasePlayerCnt = releasePlayerCnt +1
										Wend
									%>
									</table>
								</td>
							</tr>						
							<%
							 objRS.MoveNext
							 Wend
							 
							 objRS.Close
							%>
						</tbody>
					</table>
	
				</div>
				<div id="mine" class="tab-pane fade">				
					<table class="table table-responsive table-custom-black table-bordered table-condensed">
							<thead>
								<tr style="background-color:black;color:yellowgreen;font-weight:bold;">
									<th class="big" width="50%">Team</th>
									<th class="big text-center" width="5%">POS</th>											
									<th class="big text-center" width="5%">BAL</th>									
								</tr>
							</thead>
							<% Set objRSWP = Server.CreateObject("ADODB.RecordSet")	
							   objRSWP.Open "SELECT * FROM tblOwners WHERE ownerID <> 99 AND seasonOver = False " & _
							                "ORDER BY WaiverPriority ", objConn						
							%>
							<%
							 icount = 0
							 While Not objRSWP.EOF
							    icount = icount + 1
							%>
							<%if objRSWP.Fields("ownerid").Value = ownerid then %>
							<tr class="success text-center big" style="font-weight:bold;text-align:left;vertical-align:middle;">
								<td class="big"><%=objRSWP.Fields("TeamName").Value %></td>
								<td class="text-center"><%=objRSWP.Fields("WaiverPriority").Value %></td>
								<td class="big text-center">$<%=objRSWP.Fields("WaiverBal").Value %></td>								
							</tr>
							<%else %>
							<tr class="big" style="text-align:left;vertical-align:middle;background-color:white;">
								<td class="big"><%=objRSWP.Fields("TeamName").Value %></td>
								<td class="big text-center"><%=objRSWP.Fields("WaiverPriority").Value %></td>
								<td class="big text-center">$<%=objRSWP.Fields("WaiverBal").Value %></td>								
							</tr>							
							
							<%end if%>
						<%
							 objRSWP.MoveNext
							 Wend
						%>
						</table>
				</div>
				<!--WAIVER TXNS-->
				<div id="waivers" class="tab-pane fade">				
				<% 
				'objRS.Open "Select * from qryTransactions where waiverbid > 0 order by DATEVALUE(transdate) desc, waiverbid desc, transid ",objConn
				objRS.Open "Select * from qryTransactions where waiverbid > 0 order by DATEVALUE(transdate) desc, transid ",objConn
				%>

						<table class="table table-custom-black table-responsive table-bordered table-condensed" width="100%" class="display" id="example">
							<thead>
								<tr>
									<th style="width:55%;">Team | Date</th>
									<th style="width:45%;">Player</th>
								</tr>
							</thead>
							<tbody>
							<%
							While Not objRS.EOF
							addPlayerCnt = 1
							releasePlayerCnt = 1
							wTeamName = replace(UCase(objRS.Fields("shortname").Value), "THE ", "")
							'Response.Write "wTeamName  = " & wTeamName & ".<br>"
							%>
							<tr bgcolor="#FFFFFF">		
							<%if objRS.Fields("ownerID").Value = ownerid then%>
								<td class="success big" style="text-align:left;vertical-align:top;width:55%"><span class="blueTitle"><strong><%=wTeamName%></strong></span></br><%=objRS.Fields("TransDate").value%></br>
								<% if objRS.Fields("TransType").value = "Rejected " or objRS.Fields("TransType").value  = "Rental Not Played " or objRS.Fields("TransType").value = "Rejected (Roster Full) " or objRS.Fields("TransType").value = "Rejected (Bid > Balance) "  then %>
								<span class="reject"><%=objRS.Fields("TransType").value%></span>
								<%else %>
								<span class="gameTip"><%=objRS.Fields("TransType").value%></span>								
								<%end if%>
								&nbsp;<span class="dollar"><%=FormatCurrency(objRS.Fields("TransCost").value)%></span>
								
								<%if objRS.Fields("waiverbid").Value > 0 then %>
								   <%if objRS.Fields("TransType").Value = "Added "  then%>
								      </br><mark><span class="dollar">Winning Bid</span>&nbsp;<span class="dollar">- [<%= FormatCurrency(objRS.Fields("waiverbid").value)%>]</span><mark>
								   <%else%>
								      </br><span class="gameTip">Bid</span>&nbsp;<span class="dollar">- [<%= FormatCurrency(objRS.Fields("waiverbid").value)%>]</span>
								   <%end if%>	
								<%end if%>	
								</td>
							<%else%>
								<td class="big" style="text-align:left"><span class="blueTitle"><%=wTeamName%></span></br><%=objRS.Fields("TransDate").value%></br>
								<% if objRS.Fields("TransType").value = "Rejected " or objRS.Fields("TransType").value  = "Rental Not Played " or objRS.Fields("TransType").value = "Rejected (Roster Full) " or objRS.Fields("TransType").value = "Rejected (Bid > Balance) "  then %>
								<span class="reject"><%=objRS.Fields("TransType").value%></span>
								<%else %>
								<span class="gameTip"><%=objRS.Fields("TransType").value%></span>														
								<%end if%>
								&nbsp;<span class="dollar"><%=FormatCurrency(objRS.Fields("TransCost").value)%></span>
								
								<%if objRS.Fields("waiverbid").Value > 0 then%>
								   <%if objRS.Fields("TransType").Value = "Added "  then%>
								      </br><mark><span class="dollar">Winning Bid</span>&nbsp;<span class="dollar">- [<%= FormatCurrency(objRS.Fields("waiverbid").value)%>]</span></mark>
								   <%else%>
								      </br><span class="gameTip">Bid</span>&nbsp;<span class="dollar">- [<%= FormatCurrency(objRS.Fields("waiverbid").value)%>]</span>
								   <%end if%>
								<%end if%>									
								</td>
							<%end if%>	
								<td style="width:45%;">
									<table>
									<%
										While cint(addPlayerCnt) <= cint(objRS.Fields("transAddPlayerCnt").value)
									%>
										<%if cint(addPlayerCnt) = 1 then %>
											<%wRetcd = GetName(objRS.Fields("transAddPlayer1").Value,wFName,wLName)
												wTransType = left(objRS.Fields("TransType").value,6)
												'Response.Write "Transaction Type  = " & wTransType & ".<br>"
											%>
											<%if wTransType = "Waiver" or wTransType = "Reject" then %>
												<tr><td class="big" style="text-align:center;width:4%;"><span class="blackIcon"><i class="fa fa-lock" aria-hidden="true"></i></span></td><td class="big" style="text-align:left;text-decoration: line-through;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
											<%else%>
												<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
											<%end if%>
										<%elseif cint(addPlayerCnt) = 2 then %>
											<%wRetcd = GetName(objRS.Fields("transAddPlayer2").Value,wFName,wLName)%>
											<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif cint(addPlayerCnt) = 3 then%>
												<%wRetcd = GetName(objRS.Fields("transAddPlayer3").Value,wFName,wLName)%>
											<tr><td class="big" style="text-align:center;width:4%"><span class="greenIcon"><i class="fas fa-user-plus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%end if%>
									<%
										addPlayerCnt = addPlayerCnt +1
										Wend
									%>
									<%
										While releasePlayerCnt <= objRS.Fields("transReleasePlayerCnt").value
									%>								
										<%if releasePlayerCnt = 1 then %>	
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer1").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif releasePlayerCnt = 2 then %>
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer2").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%elseif releasePlayerCnt = 3 then%>
											<%wRetcd = GetName(objRS.Fields("transReleasePlayer3").Value,wFName,wLName)%>
											<tr><td style="text-align:center;width:4%"><span class="redIcon"><i class="fas fa-user-minus"></i></span></td><td class="big" style="text-align:left;"><%=wFName%>.&nbsp;<%=wLName%></td></tr>
										<%end if%>
									<%
										releasePlayerCnt = releasePlayerCnt +1
										Wend
									%>
									</table>
								</td>
							</tr>						
							<%
							 objRS.MoveNext
							 Wend
							 
							 objRS.Close
							%>
						</tbody>
					</table>

				</div>				
				<!--END OF WAIVER TXNS-->
				<div id="finance" class="tab-pane fade">
				<% 
					Set objrsMoney 	= Server.CreateObject("ADODB.RecordSet")
					Set objrsBonus 	= Server.CreateObject("ADODB.RecordSet")
					
					objrsMoney.Open  	"SELECT m.*, o.shortname,o.entry_fee,o.site_fee from QryMoney m, tblowners o " & _
														"where m.ownerid = o.ownerid order by m.totalspent desc, o.ownerid" , objConn
				
					w_total_spent = 0 
					w_entry_spent	= 0 
					w_site_spent  = 0 
					w_bonus_amt   = 25
					w_bonus1_amt  = 0 	 
					w_bonus2_amt  = 0 
					w_bonus3_amt  = 0 
					w_bonus4_amt  = 0 
					w_bonus5_amt  = 0 
					w_bonus6_amt  = 0 
					w_bonus7_amt  = 0 
					w_bonus8_amt	= 0 
					w_bonus9_amt  = 0 
					w_bonus10_amt	= 0 
					
					
				
					
				%>

            <table class="table table-striped table-bordered table-custom-black table-condensed">
						<thead>
              <tr> 
                <th style="width:25%">Name</th>
								<th style="width:25%;text-align:center">Trans</th>
								<th style="width:25%;text-align:center" >W/CYL</th>
								<th style="width:25%;text-align:center" >Bal</th>
              </tr>
						</thead>            
							<%
								 w_bonus_paid = 0
								 While Not objrsMoney.EOF
									w_bonus_amt = 0
									objrsBonus.open   "SELECT * From Standings_Cycle WHERE rank = 1 and ID = "& objrsMoney("ownerId").Value &" ",objConn,3,3,1 		
									
							
								  'Response.Write "My Bonus Record Count       = "&objrsBonus.RecordCount&".<br>"
								  'Response.Write "Owner ID Processing         = "&objrsMoney("ownerid").Value&".<br>"
									
									if objrsBonus.RecordCount > 0 then
										w_bonus_amt = objrsBonus.RecordCount * 25 
									else 	
										w_bonus_amt = 0
									end if	
									
									w_amount_due = objrsMoney("totalspent").Value - w_bonus_amt
									w_bonus_paid = w_bonus_paid + w_bonus_amt	

 									objrsBonus.Close
									
							%>
							<%if objrsMoney.Fields("ownerID").Value = ownerid then%>
								<tr class="success" style="font-weight:bold;text-align:center;vertical-align:middle;">								
							<%else%>
								<tr bgcolor="#FFFFFF">								
							<%end if%>							
                <td class="big" style="text-align:left;"><span class="blueTitle"><%= objrsMoney("shortname").Value %></span></td>
								<td class="big" style="text-align:center;"><%= FormatCurrency(objrsMoney("totalspent").Value)%></td>
								<td class="big" style="text-align:center;"><%= FormatCurrency(w_bonus_amt)%></td>
								<td class="big" style="text-align:center;"><%= FormatCurrency(w_amount_due)%></td>

							<%
							w_total_spent = w_total_spent + objrsMoney.Fields("totalspent").Value  
							w_entry_spent = w_entry_spent + objrsMoney.Fields("entry_fee").Value
							w_site_spent  = w_site_spent  + objrsMoney.Fields("site_fee").Value

							objrsMoney.MoveNext
							Wend
						 
							objrsMoney.Close
							Set objrsMoney = Nothing
							prize_money = ((w_total_spent + w_entry_spent) - (200 + 175))
							first_place = prize_money *.60
							second_place = prize_money *.30
							third_place = prize_money *.10
							%>
              </tr>
							<tr style="background-color:yellowgreen;color:black;font-weight:bold;">	
                <td class="big" style="text-align:center"><strong>Totals</strong></td>
                <td class="big" style="text-align:center"><strong><%= FormatCurrency(w_total_spent)%></strong></td>
								<td class="big" style="text-align:center"><strong><%= FormatCurrency(w_bonus_paid)%></strong></td>
								<td class="big" style="text-align:center"><strong>***</strong></td>
              </tr>
              <tr style="background-color:#FFFFFF">	
                <th class="big" style="text-align:center" colspan="6">PRIZE MONEY PAYOUT</th>
              </tr>
             <tr bgcolor="#FFFFFF">		
                <td class="big" colspan="3">Regular Season Winner</td>
                <td class="big text-center" colspan="3"><black>$200.00</black></td>
              </tr>
              <tr bgcolor="#FFFFFF">		
                <td class="big" colspan="3">IGBL Champion&nbsp;<yellowIcon><i class="fa fa-trophy"></i></yellowIcon></td>
                <td class="big text-center" colspan="3"><black><%= FormatCurrency(first_place)%></black></td>
              </tr>
              <tr bgcolor="#FFFFFF">	
                <td class="big" colspan="3">IGBL Runner-Up</td>
                <td class="big text-center" colspan="3"><black><%= FormatCurrency(second_place)%></black></td>
              </tr>
              <tr bgcolor="#FFFFFF">		
                <td class="big" colspan="3">IGBL Consolation Winner</td>
                <td class="big text-center" colspan="3"><black><%= FormatCurrency(third_place)%></black></td>
              </tr>
            </table>
				<%
					Set objrsMoney 	= Nothing
					Set objConn 		= Nothing
				%>
					</br>
            <table class="table table-bordered table-custom-black table-condensed">
						<thead>
              <tr> 
                <th class="big" style="width:75%;">Transaction Type</th>
                <th class="big" style="width:25%;text-align:center">Cost</th>
              </tr>
						</thead>            
							 <tr style="background-color:#FFFFFF">
                <td class="big">Team Entry Fee</td>
                <td class="big" style="text-align:center">$185.60</td>
              </tr>
							 <tr style="background-color:#FFFFFF">
                <td class="big">Site Hosting Fee</td>
                <td class="big" style="text-align:center">$14.40</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Free Agent Acquisition</td>
                <td class="big" style="text-align:center">$<%= wPickUp%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Rental Acquisition</td>
                <td class="big" style="text-align:center">$<%= wRental%>.00</td>
              </tr>
							<tr style="background-color:#FFFFFF">	
                <td class="big">Rental Acquisition Penalty</td>
                <td class="big" style="text-align:center">$<%= wRentPenalty%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Trade Acquisition</td>
                <td class="big" style="text-align:center">$<%= wTrade%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">
                <td class="big">Waiver Acquisition</td>
                <td class="big" style="text-align:center">$<%= wPickUp%>.00</td>
              </tr>
						</table>
						</br>	
						<strong class="red">Prize Money Distrubution:</strong></br><small>Total Transactions + Total Entry Fees - ($200 Regular Season Winner + $175 Cycle Win Shares + $144 Site Hosting Fees)</small>
						
			</div>
			<div id="txncosts" class="tab-pane fade">

            <table class="table table-bordered table-custom-black table-condensed">
						<thead>
              <tr> 
                <th class="big" style="width:75%;">Transaction Type</th>
                <th class="big" style="width:25%;text-align:center">Cost</th>
              </tr>
						</thead>            
							 <tr style="background-color:#FFFFFF">
                <td class="big">Team Entry Fee</td>
                <td class="big" style="text-align:center">$185.60</td>
              </tr>
							 <tr style="background-color:#FFFFFF">
                <td class="big">Site Hosting Fee</td>
                <td class="big" style="text-align:center">$14.40</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Free Agent Acquisition</td>
                <td class="big" style="text-align:center">$<%= wPickUp%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Rental Acquisition</td>
                <td class="big" style="text-align:center">$<%= wRental%>.00</td>
              </tr>
							<tr style="background-color:#FFFFFF">	
                <td class="big">Rental Acquisition Penalty</td>
                <td class="big" style="text-align:center">$<%= wRentPenalty%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">	
                <td class="big">Trade Acquisition</td>
                <td class="big" style="text-align:center">$<%= wTrade%>.00</td>
              </tr>
							 <tr style="background-color:#FFFFFF">
                <td class="big">Waiver Acquisition</td>
                <td class="big" style="text-align:center">$<%= wPickUp%>.00</td>
              </tr>
						</table>
						<!--<span style="font-weight:bold;"><span class="greenIcon"><i class="far fa-money-bill-alt"></i></span> Indicates Cost Increase!</span>-->

			</div>
			</div>
		</div>
	</div>
</div>
</br>
<%
 '######################################
  ' GetStats
  '   This function will be call for each player.
  '   Get First and Last Names.
  '   Get Box score information for Player.
  '   Pass information back to calling program.
  '######################################
  Function GetName (pPID,pFName,pLName)	
  
	'Response.Write "pPID = "&pPID&"<br>"
	Set objRSPlayerNames  = Server.CreateObject("ADODB.RecordSet")
	objRSPlayerNames.Open "Select firstName,lastName  from tblPlayers where PID = "& pPID , objConn,3,3,1
	'Response.Write "First Name = "&pFName&"<br>"
	'Response.Write "Last Name = "&pLName&"<br>"
	pFName  = left(objRSPlayerNames.Fields("firstName").Value,1)
	pLName  = left(objRSPlayerNames.Fields("lastName").Value,13)
	objRSPlayerNames.Close
  End Function

	%>
<%
  objRSWP.Close
	objConn.Close
  Set objConn = Nothing
%>
</body>
</html>