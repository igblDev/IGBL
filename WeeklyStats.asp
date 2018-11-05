<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCI
	Session.CodePage = 1252
	Err.Clear
	
	Dim hotStartDate,hotEndDate,objRSSort, objConn, objRSPID
	
	hotEndDate   = date()
	hotStartDate = (date() - 10)
	
	strErrorUrl = ""
	ownerid = Session("ownerid")

	if ownerid = "" then
		GetAnyParameter "var_ownerid", ownerid
	end if

	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if


	Set objConn      = Server.CreateObject("ADODB.Connection")
	Set objRSSort    = Server.CreateObject("ADODB.RecordSet")
	Set objRSPID     = Server.CreateObject("ADODB.RecordSet")
	Set objRSDate     = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"


	objRSDate.Open  "SELECT MAX(gameDate) as EndDate, " &_ 
                    "MAX(GameDate)-(select param_amount-1 from tblParameterCtl where param_name = 'WHOS_HOT_NBR_DAYS') as StartDate, " &_
				    "MIN(GameDate) as FirstGame " & _
 	                "FROM tblLast5 ", objConn,3,3,1
	
    if IsNull(objRSDate.Fields("EndDate")) then		
		hotEndDate = date()
		hotStartDate = date()
	else
	    hotEndDate = objRSDate.Fields("EndDate").Value				
		if objRSDate.Fields("FirstGame").Value > objRSDate.Fields("StartDate").Value then
	       hotStartDate = objRSDate.Fields("FirstGame").Value
		else
		   hotStartDate = objRSDate.Fields("StartDate").Value
		end if		
	end if	
	
	
	objRSDate.Close								
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
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
</script>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-fire"></i>&nbsp;Who's Hot</strong>
			</div>
		</div>
	</div>
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">	
			<span style="font-size:12px;color:black;" class="pull-right"><strong>Sortable Stats</strong></span></br>		
			<span style="font-size:12px;color:red;" class="pull-right"><strong><span style="color:black;"><%= hotStartDate %></span> to <span style="color:black;"><%= hotEndDate %></strong></span>
		</div>
	</div>
	</br>
</div>
<%
	'objRSSort.Open  "SELECT * FROM tblWeeklyStats order by BarpTot desc", objConn,3,3,1
	
	objRSSort.Open    "SELECT  first, last, count(*) as Games, " &_
                    "avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
										"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP, avg(usage) as avgUSG " &_
                    "FROM tblLast5 t " &_
                    "WHERE gamedate >= #"&hotStartDate&"# " &_
                    "group by first, last " &_
										"order by avg(BARPTot) desc ", objConn,3,3,1	    					
%>
<!--#include virtual="Common/headerMain.inc"-->

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="POST" language="JavaScript" name="FrontPage_Form1">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-bordered table-custom-black table-condensed" width="100%" class="display" id="example">
				<thead>
					<tr>
						<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
						<th class="big" style="text-decoration: underline;">Name</th>
						<th class="big" style="text-decoration: underline;">G</th>
						<th class="big hidden-xs" style="text-decoration: underline;">M</th>
						<th class="big hidden-xs" style="text-decoration: underline;">U</th>
						<th class="big" style="text-decoration: underline;">B</th>
						<th class="big" style="text-decoration: underline;">A</th>
						<th class="big" style="text-decoration: underline;">R</th>
						<th class="big" style="text-decoration: underline;">P</th>
						<th class="big" style="text-decoration: underline;">S</th>
						<th class="big" style="text-decoration: underline;">3</th>
						<th class="big" style="text-decoration: underline;">T</th>
						</tr>
				</thead>
				<tbody>
					<%
						While Not objRSSort.EOF		
						wFirstName = objRSSort.Fields("first").Value
						wLastName  = objRSSort.Fields("last").Value
												
						objRSPID.Open "SELECT p.pid, p.firstName, p.lastName,p.POS, n.teamShortName " & _
                          "FROM tblPlayers p,  tblNBATEAMS n " & _
                          "where  p.NBATeamID = n.NBATID " & _
                          "and p.firstName = '"&wFirstName&"'  " & _
                          "and p.lastName = '"&wLastName&"' ", objConn,1,1
						
						wPID      = objRSPID.Fields("PID").Value 
						wTeam     = objRSPID.Fields("teamShortName").Value
						wPOS      = objRSPID.Fields("pos").Value
						objRSPID.Close

					%>
						<tr style="background-color:white;">
							<td class="big" style="vertical-align: middle;font-weight:bold;"><%=round(objRSSort.Fields("avgBarps").Value,2)%></td>							
							<td class="big" style="vertical-align: middle;text-align:left;">								
								<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSSort.Fields("first").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("last").Value,15)%></a>&nbsp;<span class="greenTrade text-uppercase"><%=wTeam%></span>&nbsp;<span class="orange text-uppercase"><%=wPos%></span>
							</td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("Games").Value%></td>
							<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgMP").Value,0)%></td>
							<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgUSG").Value,0)%></td>									
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgBlks").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgAst").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgReb").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgPts").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgStl").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avg3pt").Value,0)%></td>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgTo").Value,0)%></td>	
						</tr>		
						<%
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
objRSPID.Close
ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>