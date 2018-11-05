<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCI
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""
	ownerid = Session("ownerid")

	if ownerid = "" then
		GetAnyParameter "var_ownerid", ownerid
	end if

	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if

	Dim objRSSort,objConn,objRSNBASked

	Set objConn      = Server.CreateObject("ADODB.Connection")
	Set objRSSort    = Server.CreateObject("ADODB.RecordSet")
	Set objRSNBASked = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

%>
	<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
p {
border: 1px solid red;
padding: 5px;
background-color:beige;
}
red {
color: red;
font-weight:bold;
}
.item {
background: #333;
text-align: center;
height: 120px !important;
}
.panel-primary {
background-color: #194719 !important;
}
td{
	text-align: center;
	vertical-align: middle;
	font-size: 11px;
}

</style>

</head>
<body>
<script language="JavaScript" type="text/javascript">
	$.fn.pageMe = function(opts){
	var $this = this,
	defaults = {
	perPage: 7,
	showPrevNext: false,
	hidePageNumbers: false
	},
	settings = $.extend(defaults, opts);

	var listElement = $this;
	var perPage = settings.perPage;
	var children = listElement.children();
	var pager = $('.pager');

	if (typeof settings.childSelector!="undefined") {
	children = listElement.find(settings.childSelector);
	}

	if (typeof settings.pagerSelector!="undefined") {
	pager = $(settings.pagerSelector);
	}

	var numItems = children.size();
	var numPages = Math.ceil(numItems/perPage);

	pager.data("curr",0);

	if (settings.showPrevNext){
	$('<li><a href="#" class="prev_link">«</a></li>').appendTo(pager);
	}

	var curr = 0;
	while(numPages > curr && (settings.hidePageNumbers==false)){
	$('<li><a href="#" class="page_link">'+(curr+1)+'</a></li>').appendTo(pager);
	curr++;
	}

	if (settings.showPrevNext){
	$('<li><a href="#" class="next_link">»</a></li>').appendTo(pager);
	}

	pager.find('.page_link:first').addClass('active');
	pager.find('.prev_link').hide();
	if (numPages<=1) {
	pager.find('.next_link').hide();
	}
	pager.children().eq(1).addClass("active");

	children.hide();
	children.slice(0, perPage).show();

	pager.find('li .page_link').click(function(){
	var clickedPage = $(this).html().valueOf()-1;
	goTo(clickedPage,perPage);
	return false;
	});
	pager.find('li .prev_link').click(function(){
	previous();
	return false;
	});
	pager.find('li .next_link').click(function(){
	next();
	return false;
	});

	function previous(){
	var goToPage = parseInt(pager.data("curr")) - 1;
	goTo(goToPage);
	}

	function next(){
	goToPage = parseInt(pager.data("curr")) + 1;
	goTo(goToPage);
	}

	function goTo(page){
	var startAt = page * perPage,
	endOn = startAt + perPage;

	children.css('display','none').slice(startAt, endOn).show();

	if (page>=1) {
	pager.find('.prev_link').show();
	}
	else {
	pager.find('.prev_link').hide();
	}

	if (page<(numPages-1)) {
	pager.find('.next_link').show();
	}
	else {
	pager.find('.next_link').hide();
	}

	pager.data("curr",page);
	pager.children().removeClass("active");
	pager.children().eq(page+1).addClass("active");

	}
	};

	$(document).ready(function(){

	$('#myTable').pageMe({pagerSelector:'#myPager',showPrevNext:true,hidePageNumbers:false,perPage:30});

	});
	
$(document).ready(function() {
    $('#example').DataTable( {
        "order": [[ 1, "desc" ]]
    } );
} );
</script>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Player Search</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-custom-black table-responsive table-bordered table-condensed">
				<tr>
					<th colspan="3" class="big">Last 5 - Player Trending</th>
				</tr>
				<tr style="background-color:white;color:black;font-weight:bold;">
					<td class="big" style="vertical-align: middle;background-color:#468847;color:white;font-weight:bold;"><i class="far fa-long-arrow-up fa-lg"></i></td>
					<td class="big" style="vertical-align: middle;background-color:#9a1400;color:white;font-weight:bold;"><i class="far fa-long-arrow-down fa-lg"></i></td>
					<td class="big" style="vertical-align: middle;background-color:gold;color:black;font-weight:bold;"><i class="fal fa-arrows-h fa-lg"></i></td>
				</tr>
				<tr style="background-color:white;color:black;font-weight:bold;font-size:13px">
					<td class="big" style="width:33%;">Rostered</td>
					<td class="big" style="width:33%;background-color:#fcf8e3;">Free Agent</td>
					<td class="big" style="width:33%;background-color:#dff0d8">My Player</td>
				</tr>
			</table>
		</div>
	</div>
</div>
</br>
<%
	objRSSort.Open  "SELECT * FROM qry_PlayerAll order by barps desc", objConn,3,3,1
%>
<!--#include virtual="Common/headerMain.inc"-->

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="POST" language="JavaScript" name="FrontPage_Form1">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example">
				<thead>
					<tr>
						<th class="big" style="width:18%;">Player</th>
						<th class="big" style="width:9%;text-decoration: underline;"><i class="fas fa-basketball-ball"></i></th>
						<th class="big" style="width:9%;text-decoration: underline;">5</th>
						<th class="big" style="width:8%;text-decoration: underline;">B</th>
						<th class="big" style="width:8%;text-decoration: underline;">A</th>
						<th class="big" style="width:8%;text-decoration: underline;">R</th>
						<th class="big" style="width:8%;text-decoration: underline;">P</th>
						<th class="big" style="width:8%;text-decoration: underline;">S</th>
						<th class="big" style="width:8%;text-decoration: underline;">3</th>
						<th class="big" style="width:8%;text-decoration: underline;">T</th>
						<th class="big" style="width:8%;text-decoration: underline;">D</th>
					</tr>
				</thead>
				<tbody>
					<%
					While Not objRSSort.EOF
					 
					 'Response.Write "DO I PLAY TODAY  " &objRSSort.Fields("NBATeamID").Value& "<br> "
					 'Response.Write "PLAYER LAST NAME  " &objRSSort.Fields("lastName").Value& "<br> "
					 
					 objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSSort.Fields("NBATeamID").Value, objConn,3,3,1
					 'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
					 if objRSNBASked.RecordCount > 0 then
					     'Response.Write "RECORD COUNT " &objRSNBASked.RecordCount& "<br> "
							 wTipTime = objRSNBASked.Fields("GameTime").Value
					 else
							 wTipTime = "12:00:00 AM"
					 end if						   
					 if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
					 else
							wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
					 end if	
					 
					
					%>
					
					<%if objRSSort.Fields("ownerID").Value > 0 then%>
						<%if objRSSort.Fields("ownerID").Value = ownerid then%>
							<tr class="success">	
						<%else%>
							<tr style="background-color:white;">		
						<%end if%>	
					<%else%>
							<tr class="warning"> 
					<%end if%>	
						<%if objRSSort.Fields("statusDesc").Value  = "On Team"then	%>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("LastName").Value,10)%></a>
								<small>
									</br><span class="redTrade  text-uppercase"><%=objRSSort.Fields("shortName").Value %></span>
									</br><span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
								</small>
							</td>
						<%else%>
								<% if objRSNBASked.RecordCount > 0 then %>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue big" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("LastName").Value,10)%></a>
								<small>
									</br><span class="redTrade  text-uppercase">free&nbsp;<i class="fas fa-clock" style="font-weight: bold;color:black;"></i></span>
									</br><span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
								</small>
							</td>
								<%else%>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue big" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("LastName").Value,10)%></a>
								<small>
									</br><span class="redTrade  text-uppercase">free</span>
									</br><span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
								</small>
							</td>
								<%end if%>		
						<%end if%>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("barps").Value,0)%></td>		
							<% if CInt(objRSSort.Fields("l5barps").Value) > CInt(objRSSort.Fields("barps").Value) then %>
							<td class="big" style="background-color:#468847;vertical-align:middle;text-align:center">
							<table class="table table-striped table-bordered table-custom-black table-condensed">
								<tr>
									<td class="big" style="vertical-align:middle;text-align:center;background-color:white;color:#468847;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
								</tr>
							</table>
						</td>
							<% elseif CInt(objRSSort.Fields("barps").Value) > CInt(objRSSort.Fields("l5barps").Value) then%>
							<td class="big" style="background-color:#9a1400;vertical-align:middle;text-align:center">
								<table class="table table-striped table-bordered table-custom-black table-condensed">
									<tr>
										<td class="big" style="vertical-align:middle;text-align:center;background-color:white;color:#9a1400;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
									</tr>
								</table>
							</td>												
							<%else%>
							<td class="big" style="background-color:gold;vertical-align:middle;text-align:center">
								<table class="table table-striped table-bordered table-custom-black table-condensed">
									<tr>
										<td class="big" style="vertical-align:middle;text-align:center;background-color:white;color:black;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
									</tr>
								</table>
							</td>	
							<%end if %>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("blk").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("ast").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("reb").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("ppg").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("stl").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("three").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("to").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("numTdbls").Value%></td>
						</tr>		
						<%
						objRSNBASked.Close
						objRSSort.MoveNext						
						Wend
						%>
				</tbody>	
				</table>
				<br>
			</div>
		</div>
	</div>
</form>
<%
objRSNBASked.Close
ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>