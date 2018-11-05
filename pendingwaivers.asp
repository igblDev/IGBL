<!-- #include file="adovbs.inc" -->
<!--#include virtual="/Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConn, objRSwaivers, ownerid, objRS,sAction,strSQL, iwaiveId, objrsCurrentWaiver,iPlayerWaived,pidSplit 
	
	
	GetAnyParameter "Action", sAction
	GetAnyParameter "var_wavierid", swaverid

	Set objConn            = Server.CreateObject("ADODB.Connection")
	Set objRSwaivers       = Server.CreateObject("ADODB.RecordSet")
	Set objRS              = Server.CreateObject("ADODB.RecordSet")
	Set objrsCurrentWaiver = Server.CreateObject("ADODB.RecordSet")
	Set objrsMoreWaivers   = Server.CreateObject("ADODB.RecordSet")
	Set objNextRundate     = Server.CreateObject("ADODB.RecordSet")

	bid_loop_cnt = 1
	objConn.Open Application("lineupstest_ConnectionString")


	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"
	%>
	<!--#include virtual="Common/session.inc"-->
	<% 	
	if sAction = "Reorder Pending Waivers" or sAction = "" then
		'DO NOTHING
	else
		pidSplit = Split(Request.Form("Action"), ";")
		sAction = pidSplit(0)
	end if
 	
	objNextRundate.Open	"SELECT * FROM tblTimedEvents where event = 'pendingwaiversall'", objConn
	nextrundate = objNextRundate.Fields("nextrun_est").Value - 1/24
	objNextRundate.Close
	
	select case sAction

	case "Reorder Pending Waivers"
		
 	    dim w_loop_count
			w_loop_count = Request.Form("pidClaimed").count
		
			'#########################################
			'# Retrieve Values from the form		    '#
			'#########################################
		
		    
			For I = 1 To w_loop_count

				if I = 1 then

					'#########################################
					'# Delete from the Waivers Table		'#
					'#########################################
					strSQL = "DELETE from tblWaivers where ownerid = " & ownerid & ";"
					objConn.Execute strSQL
					
					'#########################################
					'# Delete from the Waivers Log Table	'#
					'#########################################
				
					strSQL = "DELETE from tblwaiverlog " & _
							     "where ownerid  = " & ownerid & "  " & _
					  	     "and   created_dttm >= date() - 12/24 ; "
					objConn.Execute strSQL				
				end if
				
				ipidClaimed = Split(Request.Form("pidClaimed")(I), ";")
				ipidClaimed(0)  'The Player Claimed off Waivers	
				ipidClaimed(1)  'The Wavier Transaction ID
				ipidClaimed(2)  'The Player Waived
				ipidClaimed(3)  'The Bid Amount Waived

			
				'#########################################
				'# Insert into the Waivers Table    		'#
				'#########################################
										
				strSQL ="insert into tblwaivers(OwnerId,PID_Claimed,PID_Waived,waiverbid) " &_
                        "values ("&ownerid&","&ipidClaimed(0)&","&ipidClaimed(2)&","&ipidClaimed(3)&")"
				
				objConn.Execute strSQL
				
				strSQL ="insert into tblwaiverlog(OwnerId,PID_Claimed,PID_Waived,waiverbid) " &_
				        "values ("&ownerid&","&ipidClaimed(0)&","&ipidClaimed(2)&","&ipidClaimed(3)&")"
				objConn.Execute strSQL
				
			next
			
		case "Update Bid"	
	
			'#########################################
			'# Setup Waiver to Update/Delete 		    '#
			'#########################################	
			iwaiverID = Split(Request.Form("Action"), ";")

		case "Process Bid"	
			
			'#########################################
			'# Process Waiver Bid Amount    		    '#
			'#########################################

			ProcessWaiverID  = Request.Form("waiverID")
			newBid           = Request.Form("bidAmount") 
		
			strSQL = "UPDATE tblWaivers SET waiverbid  = "&newBid&" Where waiverId = "&ProcessWaiverID&" "
			sAction = ""
			objConn.Execute strSQL				
		

		case "Delete Pending Waivers"	
		
			'#########################################
			'# Delete from the Waivers Table		    '#
			'#########################################

			strSQL = "DELETE from tblWaivers where waiverID  =" & pidSplit(2) & ";"
			objConn.Execute strSQL
				
			'#########################################
			'# Delete from the Waivers Log Table	'#
			'#########################################
			
			strSQL = "DELETE from tblwaiverlog " & _
								 "where ownerid  = " & iOwnerid & "  " & _
								 "and   waiverID = " & pidSplit(2) & "  " & _
								 "and   created_dttm >= date() - 12/24 ; "
			objConn.Execute strSQL	

			strSQL = "UPDATE tblPlayers SET pendingwaiver = 0 "
			objConn.Execute strSQL
			
			strSQL = "UPDATE tblplayers t set t.pendingwaiver = 1 " & _
					     "WHERE exists (SELECT 1 FROM tblWaivers w where w.pid_waived = t.pid) "
			objConn.Execute strSQL			

		case ""
			'First time to the screen
			ownerid = Request.querystring("ownerid")
			if ownerid  = "" then	
				GetAnyParameter "var_ownerid", ownerid
			end if	

	end select
	objRSwaivers.Open	"SELECT * FROM qrywaivercnt WHERE (((qrywaivercnt.OwnerID)=" & ownerid & "))", 	objConn,3,3,1
%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2018-2019</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
</head>
<style>
.bs-callout-success {
	border-left-color: #354478;
	padding: 10px;
	border-left-width: 4px;
	border-radius: 3px;
	background-color: white;
}
.bs-callout-success h4 {
    color: #354478;
}

greenText {
	color:#468847;
	font-weight: 400;
  text-transform: capitalize;
}
blackText {
	color:black;
	font-weight: 400;
	text-transform: none;
}
redText {
	color:#9a1400;
	font-weight: 400;
	text-transform: capitalize;
}
.panel-title {
    color: yellowgreen;
	text-transform: none;	
	font-size: 14px  !important;
}

.bs-callout-success {
    border-left-color: #000000;
    padding: 10px;
    border-left-width: 4px;
    border-radius: 3px;
    background-color: white;
}

.panel-heading {
    background-image: none;
    background-color: #000000  !important;
    color: white;
    height: 30px;
    padding: 5px 5px;
}
a.blue {
    color: #354478 !important;
    text-decoration: underline;
}
.btn-link {
	font-weight: 400;
	color: #354478 !important;
	text-decoration: underline;
	border-radius: 0;
	font-size: 11px;
}

</style>
<body>
<script language="JavaScript" type="text/javascript"><!--
$(document).ready(function(){
    $(".up,.down").click(function(){
        var row = $(this).parents("tr:first");
        if ($(this).is(".up")) {
            row.insertBefore(row.prev());
        } else {
            row.insertAfter(row.next());
        }
    });
});
function toggle(source) {
  checkboxes = document.getElementsByName('pidClaimed');
  for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = source.checked;
  }
}
$(document).ready(function() {
		$('input[type="checkbox"]').click(function(e) {
			e.preventDefault();
			e.stopPropagation();
		});
});

window.onload = function () {
  var input = document.getElementById('myTextInput');
  input.focus();
  input.select();
}
// Sortable rows
$('.sorted_table').sortable({
  containerSelector: 'table',
  itemPath: '> tbody',
  itemSelector: 'tr',
  placeholder: '<tr class="placeholder"/>'
});

// Sortable column heads
var oldIndex;
$('.sorted_head tr').sortable({
  containerSelector: 'tr',
  itemSelector: 'th',
  placeholder: '<th class="placeholder"/>',
  vertical: false,
  onDragStart: function ($item, container, _super) {
    oldIndex = $item.index();
    $item.appendTo($item.parent());
    _super($item, container);
  },
  onDrop: function  ($item, container, _super) {
    var field,
        newIndex = $item.index();

    if(newIndex != oldIndex) {
      $item.closest('table').find('tbody tr').each(function (i, row) {
        row = $(row);
        if(newIndex < oldIndex) {
          row.children().eq(newIndex).before(row.children()[oldIndex]);
        } else if (newIndex > oldIndex) {
          row.children().eq(newIndex).after(row.children()[oldIndex]);
        }
      });
    }

    _super($item, container);
  }
});
//--></script>
<form action="pendingwaivers.asp" method="post">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="txtOwnerID" value="<%= ownerid %>" />
  <input type="hidden" name="var_wavierid" value="<%= objRSwaivers("waiverId").Value%>" />
	<input type="hidden" name="waiverID" value="<%=iwaiverID(2)%>" />
  <!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Pending Waivers	</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="bs-callout bs-callout-success">
		<h4><span style="color:#9a1400;font-weight:bold;">Handling Pending Waivers</span></h4>
			<ol>
					<li class="big"><span style="text-decoration:underline;font-weight:bold;">Waivers Run:</span>&nbsp;<span style="background-color: yellow;"><span style="font-weight:bold;"><mark><%=(FormatDateTime(nextrundate)) %> cst</mark></span></li>
					<li>Click <i class="fas fa-trash-alt"></i> Button to Delete Pending Waivers</li>
					<li>Click <i class="fas fa-edit"></i>&nbsp;Button to Update Waiver Bid Amount
					<li>Click <i class="far fa-level-up-alt fa-md red"></i> or <i class="far fa-level-down-alt fa-md red"></i> to Sort Waivers</li>
					<li>Click the Reorder Button to Execute</li>
			</ol>
	</div>
</div>
<br>
<% if sAction = "Delete Pending Waivers" then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-success">
			 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong>Success!</strong> Pending Waiver Deleted.
				</div>	
			</div>
		</div>
	</div>
<%end if%>
<% if sAction = "Reorder Pending Waivers" or sAction = "Process Bid" then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-success">
			 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong>Success!</strong> Pending Waivers Updated. 
			</div>
		</div>
	</div>
</div>
<%end if%>
<% if (sAction = "Reorder Pending Waivers" or sAction = "" or sAction = "Delete Pending Waivers") and objRSwaivers.RecordCount > 0 then %>  
<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-override">
					<table class="table table-custom-black table-bordered table-condensed">
						<tr style="background-color:black;"><td colspan="5" style="color:yellowgreen;font-weight:bold;">Current Waiver Balance: <span style="color:white;"><%= FormatCurrency(w_WaiverBal)%></span></td></tr>
						<!--<tr style="background-color:white;">
							<th style="vertical-align:middle;text-align:center;width:10%">Edit</th>
							<th style="vertical-align:middle;text-align:center;width:10%"><i class="fas fa-basketball-ball"></i></th>
							<th style="vertical-align: middle;width:60%">Players</th>
							<th style="vertical-align: middle;width:10%">Bid</th>							
							<th style="vertical-align: middle;text-align:center;width:10%">Sort</th>
						</tr>-->
						</table>
						<table class="table table-custom-black table-bordered table-condensed">
						<%
							 While Not objRSwaivers.EOF
						%>
						<tr style="background-color:white;">							
							<td style="vertical-align: middle;text-align:center;">
								<button type="submit" class="btn btn-default" value="Update Bid;<%=objRSwaivers.Fields("Pid_Claimed").Value & ";" & objRSwaivers.Fields("waiverId").Value & ";" & objRSwaivers.Fields("PID_WAIVED").Value%>" name="Action" ><i class="fas fa-edit"></i></button>						
							</td>
							<td style="vertical-align:middle;text-align:center"><input readonly checked type="checkbox" name="pidClaimed" value="<%=objRSwaivers.Fields("Pid_Claimed").Value & ";" & objRSwaivers.Fields("waiverId").Value & ";" & objRSwaivers.Fields("PID_WAIVED").Value & ";" & objRSwaivers.Fields("waiverbid").Value%>"></td>
							<td class="vertical-align: middle;">
								<greenText><i class="fa fa-plus-circle " aria-hidden="true"></i></greenText>&nbsp;<%=left(objRSwaivers("firstname").Value,14)%>&nbsp;<%=left(objRSwaivers("lastname").Value,14)%></br>
							<%if objRSwaivers("PID_WAIVED").Value > 0 then  %>
								<redText><i class="fa fa-minus-circle " aria-hidden="true"></i></redText>&nbsp;<%=left(objRSwaivers("waiverfirst").Value,14)%>&nbsp;<%=left(objRSwaivers("waiverlast").Value,14)%>
							<%else%>
								<redText><i class="fa fa-minus-circle " aria-hidden="true"></i></redText>&nbsp;<blackText>Open Roster Spot<blackText>
							<%end if%>
							</td>
							<td style="vertical-align: middle;text-align:center;"><%= FormatCurrency(objRSwaivers("waiverbid").Value)%></td>	
							<%if objRSwaivers.RecordCount > 1 then %>
							<td style="vertical-align: middle;text-align:center;">
								<a href="#" class="up blue"><i class="far fa-level-up-alt fa-md red"></i></a>&nbsp;<span><i class="far fa-ellipsis-v"></i></span>&nbsp;<a href="#" class="down blue"><i class="far fa-level-down-alt fa-md red"></i></a>
							</td>
							<%else %>
							<td style="vertical-align: middle;text-align:center;">N/A</td>							
							<%end if%>
						</tr>
						<%
							objRSwaivers.MoveNext
							Wend
						%>
					</table>
      </div>
    </div>
	</div>
</div>
<%if objRSwaivers.RecordCount > 1 then %>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<button type="submit" value="Reorder Pending Waivers" name="Action" class="btn btn-default btn-block "><i class="far fa-sort-numeric-down"></i>&nbsp;Re-Order Waiver Priority</button>
		</div>
	</div>
</div>
<%end if%>
<% elseif sAction = "Update Bid" then 
Set objrsCurrentWaiver = Server.CreateObject("ADODB.RecordSet")
objrsCurrentWaiver.Open	"SELECT * FROM qrywaivercnt WHERE qrywaivercnt.WaiverID = "&iwaiverID(2)&" ",objConn,3,3,1
%>
<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-override">
					<table class="table table-custom-black table-bordered table-condensed">
						<tr style="background-color:white;">
							<th style="vertical-align:middle;text-align:center;width:10%">Del</th>
							<th style="vertical-align: middle;width:70%">Players</th>
							<th style="vertical-align: middle;width:20%">Bid</th>							
						</tr>
						<%							
							 While Not objrsCurrentWaiver.EOF
						%>
						<tr style="background-color:white;">
							<td style="vertical-align:middle;text-align:center;">
								<button type="submit" value="Delete Pending Waivers;<%=objrsCurrentWaiver.Fields("Pid_Claimed").Value & ";" & objrsCurrentWaiver.Fields("waiverId").Value & ";" & objrsCurrentWaiver.Fields("PID_WAIVED").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fas fa-trash-alt"></i></button>						
							</td>
							<td class="vertical-align: middle;">
								<greenText><i class="fa fa-plus-circle " aria-hidden="true"></i></greenText>&nbsp;<%=left(objrsCurrentWaiver("firstname").Value,14)%>&nbsp;<%=left(objrsCurrentWaiver("lastname").Value,14)%></br>
							<%if objrsCurrentWaiver("PID_WAIVED").Value > 0 then  %>
								<redText><i class="fa fa-minus-circle " aria-hidden="true"></i></redText>&nbsp;<%=left(objrsCurrentWaiver("waiverfirst").Value,14)%>&nbsp;<%=left(objrsCurrentWaiver("waiverlast").Value,14)%>
							<%else%>
								<redText><i class="fa fa-minus-circle " aria-hidden="true"></i></redText>&nbsp;<blackText>Open Roster Spot<blackText>
							<%end if%>
							</td>
							<td style="vertical-align: middle;text-align:center;"><input class="form-control required" type="number" id="myTextInput"  name="bidAmount" value="<%=objrsCurrentWaiver("waiverbid").Value%>" autofocus  min="1" max="<%=w_WaiverBal%>" maxlength="3" size ="3" style="vertical-align: middle;text-align:center;"></td>								
						</tr>
						<%
							objrsCurrentWaiver.MoveNext
							Wend
						%>
					</table>
      </div>
    </div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<!--<button type="button" class="btn btn-default btn-xs" data-toggle="modal" data-target="#processBid"><span class="glyphicon glyphicon-save"></span>&nbsp;Update</button>-->
			<button type="submit" value="Process Bid;<%=objrsCurrentWaiver("waiverbid").Value%>" name="Action" class="btn btn-default"><span class="glyphicon glyphicon-save"></span>&nbsp;Update</button>			
			<button type="reset" value="Reset" name="Reset" class="btn btn-default"><span class="glyphicon glyphicon-refresh"></span>&nbsp;Refresh</button>
		</div>
	</div>
</div>
<%else%>
	<%
		sURL = "dashboard.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL
	%>
<% end if %>
</form>
<!--MODAL WAIVERS-->
<div class="container">
	<div class="row">
	<div class="col-md-12 col-sm-12 col-xs-12">
		<div id="processBid" class="modal fade" role="dialog">
			<div class="modal-dialog" role="document">
							
				<div class="modal-content">
					<div class="modal-header modal-header-modal">
						<button type="button" class="close" data-dismiss="modal">&times;</button>
						<h3 class="modal-title">Confirm</h3>
					</div>
					<div class="modal-body">
					<form action="pendingwaivers.asp" name="frmPlayer" method="POST">
						<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
						<input type="hidden" name="waiverID" value="<%=iwaiverID(2)%>" />
						<input type="hidden" name="bidAmount" value="<%=objrsCurrentWaiver("waiverbid").Value%>" />
						
							Are You Sure You Want to Update the Bid?</td>

					</div>		
					<div class="modal-footer">
						<button type="submit" value="Process Bid;<%=bidAmount%>" name="Action" class="btn btn-default btn-block btn-md"><i class="fa fa-road" aria-hidden="true"></i>&nbsp;Update Bid</button>
					</div>
					</div>
					</div>
					</form>
				</div>
			</div>
		</div>
	</div>
</br>
<%
  objRSwaivers.Close
  objRS.Close
  Set objRSwaivers = Nothing
  Set objRS = Nothing
  Set objConn = Nothing 
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>
