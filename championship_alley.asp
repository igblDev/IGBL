<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
td {
		
		background-color:white;
}
th {
    vertical-align:middle;
		color:black;
		
}
td.dennis {
    vertical-align:middle;
		color:green;
		
		
}
td.pat {
    vertical-align:middle;
		color:darkorange;
		
		
}
span.pat {
    vertical-align:middle;
		color:darkorange;
		
		
}
td.gary {
    vertical-align:middle;		
		color:SteelBlue;
		
		
}
td.jeff {
    vertical-align:middle;
		color:purple;
		
		
}
span.jeff {
    vertical-align:middle;
		color:purple;
		
		
}
span.keith {
    vertical-align:middle;
		color:orangered;
		
		
}
td.keith {
    vertical-align:middle;
		color:orangered;
		
		
}
td.david {
    vertical-align:middle;
		color:DarkOliveGreen;
		
		
}
td.craig {
    vertical-align:middle;
		color:red;
		
		
}
td.fred {
    vertical-align:middle;		
		color:blue;
		
		
}
td.cj {
    vertical-align:middle;
		color:DarkSlateGray;
		
		
}
span.fred {
    vertical-align:middle;
		color:blue;
		
		
}
span.cj {
    vertical-align:middle;
		color:DarkSlateGray;
		
		
}
.table-bordered>tbody>tr>td, .table-bordered>tbody>tr>th, .table-bordered>tfoot>tr>td, .table-bordered>tfoot>tr>th, .table-bordered>thead>tr>td, .table-bordered>thead>tr>th {
    border: 1px solid black;
		
}
</style>
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

 	Dim objRSgames,objRS,objConn, objRSwaivers, ownerid
	Dim strSQL, iPlayerClaimed,objRSTxns, objRSOwners, objRejectWaivers, iPlayerWaived, iOwner, w_action
	
	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
  objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
	%>
<!--#include virtual="Common/session.inc"-->
<!--#include virtual="Common/headerMain.inc"-->
</head>
<body> 
<div class="container-fluid">
  <div class="row">
    <div class="col-md-12 clearfix text-top">
      <center>
        <img class="img-responsive img-rounded" src="http://www.brooklineconnection.com/history/RecCenter/images/ChampionsLogo.JPG">
      </center>
    </div>
  </div>
	<br>
</div>
  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-custom-black table-bordered table-condensed">
					<thead>
						<tr>
							<th  class="big"   style="text-align:left" width="20%">YEAR</th>
							<th  class="big"   style="width:40%">WINNER</th>
							<th  class="big"   style="width:40%">RUNNER UP</th>
						</tr>	
					</thead>	
						<tr>
							<td class="big" style="width:20%">2017-2018</td>
							<td class="big"  style="width:40%">Jack "POT" White</td>
							<td class="big"  style="width:40%">Jeff Peskin/Keith Dlott</td>
						</tr>					
						<tr>
							<td class="big"  style="width:20%">2016-2017</td>
							<td class="big" style="width:40%">Tone Adams/<span class="pat big">Pat Harris</span></td>
							<td class="big"  style="width:40%">Dennis Myers</td>
						</tr>
						<tr>
							<td class="big"  style="width:20%">2015-2016</td>
							<td style="width:40%"><span class="jeff big">Jeff Peskin</span>/<span class="keith big">Keith Dlott</span></td>
							<td class="big" style="width:40%">Dennis Myers</td>
						</tr>
						<tr>
							<td class="big"  style="width:20%">2014-2015</td>
							<td class="jeff big" style="width:40%">Jeff Peskin</td>
							<td class="big"  style="width:40%">Tone Adams/Pat Harris</td>
						</tr>
						<tr>
							<td class="big"  style="width:20%">2013-2014</td>
							<td class="david big" style="width:40%">David Babineaux</td>
							<td class="big"  style="width:40%">Dennis Myers</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2012-2013</td>
							<td class="pat big" style="width:40%">Patrick Harris</td>
							<td class="big" style="width:40%">Jeff Peskin</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2011-2012</td>
							<td class="jeff big" style="width:40%">Jeff Peskin</td>
							<td class="big" style="width:40%">David Babineaux</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2010-2011</td>
							<td style="width:40%"><span class="fred big">Fred Curry</span>/<span class="cj big">Chris Jones</span></td>
							<td class="big" style="width:40%">Gary Rothballer</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2009-2010</td>
							<td class="cj big" style="width:40%">Chris Jones</td>
							<td class="big" style="width:40%">Patrick Harris</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2008-2009</td>
							<td class="pat big" style="width:40%">Patrick Harris</td>
							<td class="big" style="width:40%">Jack White</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2007-2008</td>
							<td class="dennis big" style="width:40%">Dennis Myers</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2006-2007</td>
							<td class="david big" style="width:40%">David Babineaux</td>
							<td class="big" style="width:40%">Jeff Peskin</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2005-2006</td>
							<td class="gary big" style="width:40%">Gary Rothballer</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2004-2005</td>
							<td class="dennis big" style="width:40%">Dennis Myers</td>
							<td class="big" style="width:40%">Chris Jones</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2003-2004</td>
							<td class="big" style="width:40%">Cliff Fox</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2002-2003</td>
							<td class="fred big" style="width:40%">Fred Curry</td>
							<td class="big" style="width:40%">Keith Dlott</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2001-2002</td>
							<td class="keith big">Keith Dlott</td>
							<td class="big" style="width:40%">Cliff Fox</td>
						</tr>
						<tr>
							<td class="big" style="width:20%">2000-2001</td>
							<td class="big" style="width:40%">Anthony White</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1999-2000</td>
							<td class="dennis big" style="width:40%">Dennis Myers</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1998-1999</td>
							<td class="dennis big" style="width:40%">Dennis Myers</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1997-1998</td>
							<td class="dennis big" style="width:40%">Dennis Myers</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1996-1997</td>
							<td class="gary big" style="width:40%">Gary Rothballer</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1995-1996</td>
							<td class="big" style="width:40%">Jeff Campbell</td>
							<td class="big" style="width:40%"></td>
						</tr>
						<tr>
							<td class="big" style="width:20%">1994-1995</td>
							<td class="big" style="width:40%">Craig Sklar</td>
							<td class="big" style="width:40%">Dennis Myers</td>
						</tr>
				</table>
			</div>
		</div>
	</div>
</body>
</html>