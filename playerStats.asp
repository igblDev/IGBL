<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	strErrorUrl = ""

	Dim objConn, sTeam, sAction, sURL,ownerid


	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%	
	ownerid = session("ownerid")
						
	%>
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<!--#include virtual="Common/headerMain.inc"-->
<style>
darkorange {
	color: darkorange;
}

th {
    vertical-align: middle;
  	text-align: center;
}

td{
	text-align: center;
	vertical-align: middle;
	font-size: 11px;
}

black {
	color:black;
	text-transform: uppercase;
}
.nav-tabs {
    border-bottom: 2px solid black;
}

</style>
</head>
<body>
<script>
$(document).ready(function() {
    $('#example1').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
$(document).ready(function() {
    $('#weeks1').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
$(document).ready(function() {
    $('#weeks2').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
$(document).ready(function() {
    $('#example2').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
</script>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-users"></i>&nbsp;Player Stats</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#Season"><i class="far fa-list-ol"></i>&nbsp;Season</a></li>
				<li><a data-toggle="tab" href="#lastNight"><i class="fal fa-sort-amount-down"></i>&nbsp;Nightly Leaders</a></li>
				<li><a data-toggle="tab" href="#1Week"><i class="fas fa-fire red"></i>&nbsp;Last 7 Days</a></li>
			</ul>
		</div>
	</div>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<!--LAST NIGHT STATS-->
				<div id="lastNight" class="tab-pane">
					<!--#include virtual="Common/leaders.inc"-->
				</div>
				<!--1 WEEK STATS-->
				<div id="1Week" class="tab-pane">
					<!--#include virtual="Common/weekly.inc"-->
				</div>
				<!--SEASON STATS-->
				<div id="Season" class="tab-pane fade in active">
					<!--#include virtual="Common/players.inc"-->
				</div>
			</div>	
		</div>		
	</div>
</div>	
<% 

Set objrs = Nothing
Set objConn = Nothing
 %>
</body>
</html>
