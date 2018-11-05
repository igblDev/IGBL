<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="Dee M. Micheals" content="">
<title>IGBL 2015-2016</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/app_nologin.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

 	Dim objRSgames,objRS,objConn, objRSwaivers, ownerid
	Dim strSQL, iPlayerClaimed,objRSTxns, objRSOwners, objRejectWaivers, iPlayerWaived, iOwner, w_action

	ownerid = session("ownerid")	
	
  if ownerid = "" then
    GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	
	
	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRS         = Server.CreateObject("ADODB.RecordSet")
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")


   objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                 "Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

	objRSgames.Open "qryStangings", objConn
	
	


%>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<base target="_self">
<style>
p {
	border: 1px solid red;
	padding: 5px;
	background-color:#FFFFCC;
}
green {
	color: green;
}
yellow {
	color: yellow;
}
red {
	color: red;
}
white {
	color: white;
}
th {
    vertical-align: middle;
  	text-align: center;
		color: DarkOrange  ;
}
td {
    vertical-align: middle;
    text-align: center;
}
black {
		color: black;
}
orange {
		color: darkorange;
		font-weight: bold;
}


.badge {
    display: inline-block;
    min-width: 10px;
    padding: 3px 12px;
    font-size: 12px;
    font-weight: 1200;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: #4688412;
    border-radius: 10px;
}
.bs-callout-success {
	border-left-color: darkorange;
	padding: 10px;
	border-left-width: 5px;
	border-radius: 3px;
	background-color: white;
}
.bs-callout-success h4 {
    color: darkorange;
}
.modal-header-success {
    color:#fff;
    padding:9px 15px;
    border-bottom:1px solid #eee;
    background-color: darkorange;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
     border-top-left-radius: 5px;
     border-top-right-radius: 5px;
}
.navbar-custom {
	background-color:darkorange;
    color:#ffffff;
  	border-radius:0;
}
  
.navbar-custom .navbar-nav > li > a {
  	color:#fff;
}
.navbar-custom .navbar-nav > .active > a, .navbar-nav > .active > a:hover, .navbar-nav > .active > a:focus {
    color: #ffffff;
	background-color:transparent;
}
      
.navbar-custom .navbar-nav > li > a:hover, .nav > li > a:focus {
    text-decoration: none;
    background-color: orange;
}
      
.navbar-custom .navbar-brand {
  	color:#eeeeee;
}
.navbar-custom .navbar-toggle {
  	background-color:#eeeeee;
}
.navbar-custom .icon-bar {
  	background-color:#33aa33;
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
    $('#example').DataTable( {
        "lengthMenu": [[15, 25, 50, -1], [15, 25, 50, "All"]]
    } );
	} );
</script>
<!--#include virtual="Common/headerMainStandings.inc"-->
<!-- Trigger the modal with a button -->

<br>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<button type="button" class="btn btn-default btn-sm" data-toggle="modal" data-target="#myModal">View Last 5</button>
		</div>
	</div>
</div>
<br>
<div class="container">
	<div class="row">
		<div id="myModal" class="modal fade" role="dialog">
		<div class="modal-dialog">
    <!-- Modal content-->
    <div class="modal-content">
      <div class="modal-header modal-header-success">
        <button type="button" class="close" data-dismiss="modal">&times;</button>
        <h5 class="modal-title">Last 5</h5>
      </div>
			<div class="col-md-12 col-sm-12 col-xs-12">
			<br>
			<table class="table table-condensed">
				<thead>
					<tr style="background-color:#4688412;color:white">
						<th  nowrap  width="12%"><small><black><strong>blk</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>ast</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>reb</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>pts</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>stl</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>3pt</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>tos</strong></black></small></th>
						<th  nowrap  width="12%"><small><black><strong>bps</strong></black></small></th>
					</tr>
				</thead>
				<tr>
					<td  nowrap><small>1</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>16</small></td>
					<td  nowrap><small>26</small></td>
					<td  nowrap><small>2</small></td>
					<td  nowrap><small>0</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>45</small></td>
				</tr>
				<tr>
					<td  nowrap><small>1</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>16</small></td>
					<td  nowrap><small>26</small></td>
					<td  nowrap><small>2</small></td>
					<td  nowrap><small>0</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>45</small></td>
				</tr>
				<tr>
					<td  nowrap><small>1</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>16</small></td>
					<td  nowrap><small>26</small></td>
					<td  nowrap><small>2</small></td>
					<td  nowrap><small>0</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>45</small></td>
				</tr>
				<tr>
					<td  nowrap><small>1</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>16</small></td>
					<td  nowrap><small>26</small></td>
					<td  nowrap><small>2</small></td>
					<td  nowrap><small>0</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>45</small></td>
				</tr>
				<tr>
					<td  nowrap><small>1</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>16</small></td>
					<td  nowrap><small>26</small></td>
					<td  nowrap><small>2</small></td>
					<td  nowrap><small>0</small></td>
					<td  nowrap><small>3</small></td>
					<td  nowrap><small>45</small></td>
				</tr>
				<tr>
					<td colspan="8"><small><strong>5 Game Average</strong></small></td>
				</tr>
				<tr>
					<td nowrap><small><orange>1</orange></small></td>
					<td nowrap><small><orange>3</orange></small></td>
					<td nowrap><small><orange>16</orange></small></td>
					<td nowrap><small><orange>26</orange></small></td>
					<td nowrap><small><orange>2</orange></small></td>
					<td nowrap><small><orange>0</orange></small></td>
					<td nowrap><small><orange>3</orange></small></td>
					<td nowrap><small><orange>45</orange></small></td>
				</tr>
				</table>
				<br>
    </div>
    <div class="modal-footer">
			<button type="button" class=" btn btn-sm btn-default" data-dismiss="modal">Close</button>
		</div>
    </div>
		</div>
		</div>
	</div>
</div>
<br>
<%
objrs.close
objrsgames.close
objconn.close
Set objrs = Nothing
Set objrsgames = Nothing
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>
