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

	Dim objRSAuction,objConn

	Set objConn      = Server.CreateObject("ADODB.Connection")
	Set objRSAuction = Server.CreateObject("ADODB.RecordSet")

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
        "lengthMenu": [[15, 30, 75, -1], [15, 30, 75, "All"]],
				 "order": [[ 1, "desc" ]]
    } );
	} );
</script>
<%
	objRSAuction.Open  "SELECT * FROM qry_auction_recap", objConn,3,3,1
%>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-usd-circle"></i>&nbsp;Auction Recap</strong>
			</div>
		</div>
	</div>
</div>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="POST" language="JavaScript" name="FrontPage_Form1">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example">
				<thead>
					<tr>
						<th>Player - TM</th>
						<th  align="center"><i class="fas fa-usd-circle"></i></th>
						<th  align="center">AB#</th>
						<!--<th  align="center">Team</th>-->
						<th  align="center">Owner</th>
					</tr>
				</thead>
				<tbody>
					<%
					While Not objRSAuction.EOF
					%>
					<tr>
						<td class="big" valign="middle"><%=left(objRSAuction.Fields("LastName").Value,14) %>&nbsp;<%=objRSAuction.Fields("Team").Value%></td>
						<td class="big" valign="middle" align="center">$<%=objRSAuction.Fields("Auction_Price").Value %></td>
						<td class="big" valign="middle" align="center"><%=objRSAuction.Fields("Auction_Number").Value %></td>
						<!--<td  valign="middle" align="center"><%=objRSAuction.Fields("TeamName").Value %></td>-->
						<td class="big" valign="middle"><%=objRSAuction.Fields("OwnerName").Value %></td>
					</tr>		
						<%
						objRSAuction.MoveNext
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

ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>